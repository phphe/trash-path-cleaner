#Requires -Version 2.0
<#
    回收站定点清理 / trash-path-cleaner（Windows 11 / 10 / 7）

    作用：把「回收站里、原始路径位于指定目录下、且原始大小小于阈值」的文件永久删除。
          可选连带删除这些目录下“整目录被删除”产生的目录条目。

    原理：直接读取各驱动器 $Recycle.Bin\<SID>\ 下的 $I 元数据文件（含原始路径 / 原始大小 /
          删除时间），删除对应的 $R 实体与 $I 元数据。不经过 Shell COM 逐项枚举，速度快。

    兼容：PowerShell 2.0 及以上（Win7 自带 2.0、Win10/11 自带 5.1）。
          $I 元数据支持 Win7/Vista 的 V1（544 字节）与 Win10 1809+ 的 V2 两种格式。

    用法：双击 run-trash-path-cleaner.bat；或在 PowerShell 中：
          .\trash-path-cleaner.ps1
          .\trash-path-cleaner.ps1 -TargetPath 'D:\proj' -MaxSizeKB 100
          .\trash-path-cleaner.ps1 -TargetPath 'D:\proj;E:\tmp'      # 多个目录用分号分隔
          .\trash-path-cleaner.ps1 -AllUsers        # 自动提权，检查所有用户的回收站
          .\trash-path-cleaner.ps1 -Lang en         # 强制英文界面（默认跟随系统 UI 语言）

    界面语言：默认按系统 UI 语言自动选择中文（zh）或英文（en），非中文系统一律用英文。
#>
[CmdletBinding()]
param(
    # 目标目录（可多个）。不传则使用下面的 $ConfigTargets
    [string[]]$TargetPath,
    # 大小阈值（KB，严格小于才删除）。不传则使用下面的 $ConfigMaxSizeKB
    [int]$MaxSizeKB = 0,
    # 界面语言：zh / en。不传则按系统 UI 语言自动选择
    [string]$Lang,
    # 以管理员身份重新运行，检查所有用户/系统账户的回收站
    [switch]$AllUsers,
    # 内部参数：表示当前已是提权后的进程，避免循环提权
    [switch]$Elevated,
    # 结束后不等待回车（便于自动化）
    [switch]$NoPause
)

# ======================= 用户设置 =======================
# 需要清理的原始目录，可写多个。回收站条目只要原始路径 == 该目录 或 位于其子目录下，就算命中。
$ConfigTargets = @(
    'C:\Users\phphe\Documents\workspace',
    'C:\Users\phphe\AppData'
)

# 大小上限（KB）：原始大小严格小于该值的文件才会被删除
$ConfigMaxSizeKB = 30
# =======================================================

$ErrorActionPreference = 'Stop'
try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch { }
$OutputEncoding = [System.Text.Encoding]::UTF8

# 脚本自身路径/目录（兼容 PowerShell 2.0：没有 $PSScriptRoot / $PSCommandPath）
$ScriptPath = $PSCommandPath
if (-not $ScriptPath) { $ScriptPath = $MyInvocation.MyCommand.Definition }
$ScriptDir = $PSScriptRoot
if (-not $ScriptDir -and $ScriptPath -and (Test-Path -LiteralPath $ScriptPath)) {
    $ScriptDir = Split-Path -Parent $ScriptPath
}
if (-not $ScriptDir) { $ScriptDir = (Get-Location).Path }

# .NET 4.0+ 才有 EnumerateFiles；Win7 自带的 PowerShell 2.0 跑在 CLR 2.0 上没有，需要降级路径
$HasEnumerateFiles = $null -ne ([IO.DirectoryInfo].GetMethod('EnumerateFiles', [type[]]@([string], [IO.SearchOption])))

# ======================= 界面语言 =======================
# 默认跟随系统 UI 语言：中文系统用中文，其他系统一律用英文；可用 -Lang zh / -Lang en 强制指定。
# 语言只在启动时判断一次，所有文案预先放进 $UI，后续输出直接取值，不在循环里重复判断。
$LangCode = ''
if (-not [string]::IsNullOrEmpty($Lang)) { $LangCode = $Lang.Trim().ToLower() }
if ($LangCode -ne 'zh' -and $LangCode -ne 'en') {
    $uiName = ''
    try { $uiName = (Get-UICulture).Name } catch { }
    if ([string]::IsNullOrEmpty($uiName)) {
        try { $uiName = [Globalization.CultureInfo]::CurrentUICulture.Name } catch { $uiName = '' }
    }
    # 少数服务/沙箱环境会把线程 UI 文化改成英文，这里再补一次系统安装语言判断
    if ($uiName -notlike 'zh*') {
        try {
            if ([Globalization.CultureInfo]::InstalledUICulture.Name -like 'zh*') { $uiName = 'zh-CN' }
        } catch { }
    }
    if ($uiName -like 'zh*') { $LangCode = 'zh' } else { $LangCode = 'en' }
}

if ($LangCode -eq 'zh') {
    $UI = @{
        Title          = '          回收站定点清理'
        Targets        = '目标目录：'
        TargetMissing  = '  (目录当前不存在，按路径前缀匹配)'
        SizeCondition  = '大小条件：原始大小 < {0} KB'
        NoTargets      = '错误：没有配置任何有效的目标目录，请编辑脚本顶部的 $ConfigTargets。'
        RequestingAdmin = '正在请求管理员权限，以检查所有用户的回收站……'
        ElevatedStarted = '已在新窗口中以管理员身份运行，本窗口可关闭。'
        ElevateFailed  = '提权失败（可能取消了 UAC），继续以当前权限运行。'
        Scanning       = '正在扫描回收站……'
        ScanDone       = '扫描完成：{0} 个 SID 回收站，读取 $I 元数据 {1} 个（无法解析 {2} 个），用时 {3:N2} 秒。'
        CountingNested = '正在统计目录条目内部的嵌套文件……'
        NestedDone     = '目录递归：{0} 个目录条目内共 {1} 个文件，用时 {2:N2} 秒{3}。'
        NestedDoneSkip = '（{0} 个子目录无权限已跳过）'
        DeniedHint1    = '提示：有 {0} 个回收站目录（属于其他用户/系统账户）需要管理员权限才能读取，已跳过。'
        DeniedHint2    = '      如需一并处理，请右键脚本所在目录的 bat 选择“以管理员身份运行”，或加参数 -AllUsers。'
        PerTarget      = '[{0}] {1}'
        FilesLine      = '     文件：将删除 {0} 个，合计 {1}；不删除 {2} 个，合计 {3}（大小 ≥ {4} KB）'
        DirsLine       = '     目录：{0} 个，合计 {1}（只有选 2 才会删除；选 1 时其中的小文件会被删除，目录本身保留）'
        TotalFilesLine = '合计：文件 将删除 {0} 个 / {1}，不删除 {2} 个 / {3}'
        TotalDirsLine  = '      目录 {0} 个 / {1}（只有选 2 才会删除）'
        NothingFound   = '没有找到符合条件的回收站条目，无需处理。'
        MenuTitle      = '请选择操作：'
        Menu1          = '  1  只删除文件：{0} 个，合计 {1}（目录保留）'
        Menu2          = '  2  文件 + 目录一起删除：文件 {0} 个 / {1}，目录 {2} 个 / {3}'
        Menu3          = '  3  取消（不做任何修改）'
        MenuPrompt     = '请输入 1 / 2 / 3 后回车（默认 3）：'
        Cancelled      = '已取消，没有删除任何内容。'
        DirWarning     = '注意：目录条目内的全部内容（包括大于 {0} KB 的文件）都会随文件夹一起永久删除。'
        ConfirmSummary = '即将永久删除：文件 {0} 个，目录条目 {1} 个。删除后无法通过回收站还原。'
        ConfirmPrompt  = '确认执行请输入 y 后回车（其他任意输入取消）：'
        DoneKeepDirs   = '删除完成：文件 {0} 个，目录已保留（{1} 个），失败 {2} 个，用时 {3:N2} 秒。'
        DoneWithDirs   = '删除完成：文件 {0} 个，目录 {1} 个，失败 {2} 个，用时 {3:N2} 秒。'
        FailHeader     = '失败明细（前 20 条）：'
        FailMore       = '  …… 其余 {0} 条略'
        FailHintAdmin  = '部分条目可能属于其他用户或需要更高权限，可尝试以管理员身份运行。'
        FailHintBusy   = '部分条目可能正被其他程序占用，关闭相关程序后重试。'
        PathGuardFail  = '路径校验失败，已跳过'
        UnhandledError = '执行出错：'
        ErrorAt        = '位置：'
        PressEnter     = '按回车键退出……'
    }
} else {
    $UI = @{
        Title          = '          trash-path-cleaner'
        Targets        = 'Target folders:'
        TargetMissing  = '  (missing now; still matched by path prefix)'
        SizeCondition  = 'Size rule: original size < {0} KB'
        NoTargets      = 'Error: no valid target folder. Edit $ConfigTargets at the top of the script.'
        RequestingAdmin = 'Requesting administrator rights to check every user''s recycle bin...'
        ElevatedStarted = 'Started in a new elevated window. You can close this one.'
        ElevateFailed  = 'Elevation failed (UAC may have been cancelled); continuing with current rights.'
        Scanning       = 'Scanning recycle bins...'
        ScanDone       = 'Scan done: {0} SID recycle bin(s), {1} $I entries read ({2} unreadable), {3:N2} s.'
        CountingNested = 'Counting files inside deleted folders...'
        NestedDone     = 'Folder scan: {1} file(s) in {0} deleted folder(s), {2:N2} s{3}.'
        NestedDoneSkip = ' ({0} subfolder(s) skipped, access denied)'
        DeniedHint1    = 'Note: {0} recycle bin folder(s) belong to other users or system accounts and were skipped.'
        DeniedHint2    = '      To include them, run the .bat as administrator, or add -AllUsers.'
        PerTarget      = '[{0}] {1}'
        FilesLine      = '     Files: to delete {0} ({1}); to keep {2} ({3}) (size >= {4} KB)'
        DirsLine       = '     Folders: {0} ({1}) (deleted only by option 2; option 1 deletes small files inside but keeps the folder)'
        TotalFilesLine = 'Total: files to delete {0} / {1}, to keep {2} / {3}'
        TotalDirsLine  = '       folders {0} / {1} (deleted only by option 2)'
        NothingFound   = 'No matching recycle bin items. Nothing to do.'
        MenuTitle      = 'Choose an action:'
        Menu1          = '  1  Delete files only: {0} ({1}) (folders kept)'
        Menu2          = '  2  Delete files and folders: files {0} / {1}, folders {2} / {3}'
        Menu3          = '  3  Cancel (no changes)'
        MenuPrompt     = 'Enter 1 / 2 / 3 and press Enter (default 3): '
        Cancelled      = 'Cancelled. Nothing was deleted.'
        DirWarning     = 'Warning: everything inside these folders (including files >= {0} KB) will be deleted with them.'
        ConfirmSummary = 'About to permanently delete {0} file(s) and {1} folder entry(ies). This cannot be undone.'
        ConfirmPrompt  = 'Type y and press Enter to confirm (anything else cancels): '
        DoneKeepDirs   = 'Done: {0} file(s) deleted, {1} folder(s) kept, {2} failed, {3:N2} s.'
        DoneWithDirs   = 'Done: {0} file(s) and {1} folder(s) deleted, {2} failed, {3:N2} s.'
        FailHeader     = 'Failures (first 20):'
        FailMore       = '  ... {0} more omitted'
        FailHintAdmin  = 'Some items may belong to other users or need higher rights. Try running as administrator.'
        FailHintBusy   = 'Some items may be locked by other programs. Close them and try again.'
        PathGuardFail  = 'path check failed, skipped'
        UnhandledError = 'Error: '
        ErrorAt        = 'At: '
        PressEnter     = 'Press Enter to exit...'
    }
}

# ---------------------------------------------------------------- 工具函数

# 判断字符串为空/全空白（PowerShell 2.0 的 .NET 3.5 没有 IsNullOrWhiteSpace）
function Test-Blank {
    param([string]$Text)
    return ($null -eq $Text -or $Text.Trim().Length -eq 0)
}

function Test-IsAdmin {
    try {
        $id = [Security.Principal.WindowsIdentity]::GetCurrent()
        return (New-Object Security.Principal.WindowsPrincipal($id)).IsInRole(
            [Security.Principal.WindowsBuiltInRole]::Administrator)
    } catch { return $false }
}

# 规范化目录：去引号/空格、展开环境变量、转绝对路径、去尾部反斜杠（盘根除外）
function Convert-ToNormalPath {
    param([string]$Path, [string]$BaseDir)
    if (Test-Blank $Path) { return $null }
    $p = $Path.Trim().Trim('"').Trim("'").Trim()
    if ($p -eq '') { return $null }
    $p = [Environment]::ExpandEnvironmentVariables($p)
    try {
        if (-not [IO.Path]::IsPathRooted($p)) { $p = Join-Path $BaseDir $p }
        $p = [IO.Path]::GetFullPath($p)
    } catch { return $null }
    if ($p -match '^[A-Za-z]:\\?$') { return $p.Substring(0, 2) + '\' }
    return $p.TrimEnd('\')
}

# 路径前缀匹配，注意边界：C:\a 不会匹配 C:\ab
function Test-PathUnder {
    param([string]$Path, [string]$Prefix)
    if ([string]::IsNullOrEmpty($Path) -or [string]::IsNullOrEmpty($Prefix)) { return $false }
    if ($Path.Length -lt $Prefix.Length) { return $false }
    if ($Path.Equals($Prefix, [StringComparison]::OrdinalIgnoreCase)) { return $true }
    if ($Prefix.EndsWith('\')) {
        return $Path.StartsWith($Prefix, [StringComparison]::OrdinalIgnoreCase)
    }
    return $Path.StartsWith($Prefix + '\', [StringComparison]::OrdinalIgnoreCase)
}

function Format-Size {
    param([int64]$Bytes)
    if ($Bytes -lt 1KB) { return "$Bytes B" }
    if ($Bytes -lt 1MB) { return ('{0:N1} KB' -f ($Bytes / 1KB)) }
    if ($Bytes -lt 1GB) { return ('{0:N1} MB' -f ($Bytes / 1MB)) }
    return ('{0:N2} GB' -f ($Bytes / 1GB))
}

# 读一行输入；stdin 结束（EOF）时返回空串，避免自动化场景卡死
function Read-UserInput {
    param([string]$Prompt)
    Write-Host $Prompt -NoNewline
    $line = [Console]::ReadLine()
    if ($null -eq $line) { Write-Host ''; return '' }
    return $line.Trim()
}

# 解析单个 $I 元数据文件。返回 $null 表示无法识别。
#   V1：8 字节版本(=1) + 8 字节原始大小 + 8 字节删除时间 + 固定 520 字节 UTF-16 路径
#   V2：8 字节版本(=2) + 8 字节原始大小 + 8 字节删除时间 + 4 字节字符数 + UTF-16 路径
function Read-IndexEntry {
    param([string]$IndexPath)
    $b = [IO.File]::ReadAllBytes($IndexPath)
    if ($b.Length -lt 24) { return $null }

    $version = [BitConverter]::ToInt64($b, 0)
    $size = [BitConverter]::ToInt64($b, 8)
    $fileTime = [BitConverter]::ToInt64($b, 16)

    if ($version -eq 2) {
        if ($b.Length -lt 28) { return $null }
        $chars = [BitConverter]::ToInt32($b, 24)
        if ($chars -le 0 -or $chars -gt 32767) { return $null }
        $byteCount = [Math]::Min($chars * 2, $b.Length - 28)
        if ($byteCount -le 0) { return $null }
        $path = [Text.Encoding]::Unicode.GetString($b, 28, $byteCount)
    } elseif ($version -eq 1) {
        if ($b.Length -le 24) { return $null }
        $path = [Text.Encoding]::Unicode.GetString($b, 24, $b.Length - 24)
    } else {
        return $null
    }

    $path = $path.TrimEnd([char]0)
    if (Test-Blank $path) { return $null }

    $deletedAt = if ($fileTime -gt 0) { [DateTime]::FromFileTime($fileTime) } else { [DateTime]::MinValue }
    return New-Object PSObject -Property @{
        OriginalPath = $path
        Size         = $size
        DeletedAt    = $deletedAt
        Version      = $version
    }
}

# 超长路径转 \\?\ 形式，避免 260 字符限制
function Convert-ToExtendedPath {
    param([string]$Path)
    if ($Path.StartsWith('\\?\')) { return $Path }
    if ($Path.Length -ge 240) { return '\\?\' + $Path }
    return $Path
}

# 清除整棵树的只读/隐藏/系统属性，便于删除
function Reset-Attributes {
    param([string]$Root)
    $p = Convert-ToExtendedPath $Root
    try {
        foreach ($f in [IO.Directory]::GetFiles($p, '*', [IO.SearchOption]::AllDirectories)) {
            try { [IO.File]::SetAttributes($f, [IO.FileAttributes]::Normal) } catch { }
        }
        foreach ($d in [IO.Directory]::GetDirectories($p, '*', [IO.SearchOption]::AllDirectories)) {
            try { [IO.File]::SetAttributes($d, [IO.FileAttributes]::Normal) } catch { }
        }
        [IO.File]::SetAttributes($p, [IO.FileAttributes]::Normal)
    } catch { }
}

# 删除 $R 实体（文件或整个目录，目录含全部内容）
function Remove-RecycleData {
    param([string]$Path)
    $p = Convert-ToExtendedPath $Path
    if ([IO.Directory]::Exists($p)) {
        try {
            [IO.Directory]::Delete($p, $true)
        } catch {
            Reset-Attributes $p
            [IO.Directory]::Delete($p, $true)
        }
    } elseif ([IO.File]::Exists($p)) {
        try {
            [IO.File]::Delete($p)
        } catch {
            [IO.File]::SetAttributes($p, [IO.FileAttributes]::Normal)
            [IO.File]::Delete($p)
        }
    }
}

# 删除单个文件（$I 元数据或目录条目内部的普通文件），自动处理只读属性
function Remove-FileForce {
    param([string]$Path)
    $p = Convert-ToExtendedPath $Path
    try {
        [IO.File]::Delete($p)
    } catch {
        [IO.File]::SetAttributes($p, [IO.FileAttributes]::Normal)
        [IO.File]::Delete($p)
    }
}

# 递归枚举某个目录条目（$R 目录）内部的全部文件，只收集大小 < 阈值的。
# 正常情况一次 EnumerateFiles 扫完；仅当遇到无权限目录时才退化为逐层遍历，跳过读不了的子目录。
function Get-SmallFilesUnder {
    param([string]$Root, [int64]$LimitBytes)
    $res = New-Object PSObject -Property @{
        Files       = New-Object 'System.Collections.Generic.List[string]'
        Bytes       = [int64]0
        Scanned     = 0
        TotalBytes  = [int64]0
        SkippedDirs = 0
    }
    $rootExt = Convert-ToExtendedPath $Root

    # 快速路径：.NET 4.0+（Win10/11 的 PowerShell 5.1）用 EnumerateFiles，游标式枚举、长度来自同一句柄
    if ($HasEnumerateFiles) {
        try {
            $di = New-Object IO.DirectoryInfo $rootExt
            foreach ($fi in $di.EnumerateFiles('*', [IO.SearchOption]::AllDirectories)) {
                $res.Scanned++
                $res.TotalBytes += $fi.Length
                if ($fi.Length -lt $LimitBytes) {
                    $p = $fi.FullName
                    if ($p.StartsWith('\\?\')) { $p = $p.Substring(4) }
                    [void]$res.Files.Add($p)
                    $res.Bytes += $fi.Length
                }
            }
            return $res
        } catch {
            # 有子目录读不了时，丢弃半截结果，走下面的逐层遍历
            $res.Files.Clear()
            $res.Bytes = [int64]0
            $res.Scanned = 0
            $res.TotalBytes = [int64]0
        }
    } else {
        # .NET 2.0/3.5（Win7 自带 PowerShell 2.0）没有 EnumerateFiles，用一次 GetFiles 取回全部路径
        try {
            foreach ($p0 in [IO.Directory]::GetFiles($rootExt, '*', [IO.SearchOption]::AllDirectories)) {
                $len = [int64](New-Object IO.FileInfo $p0).Length
                $res.Scanned++
                $res.TotalBytes += $len
                if ($len -lt $LimitBytes) {
                    $p = $p0
                    if ($p.StartsWith('\\?\')) { $p = $p.Substring(4) }
                    [void]$res.Files.Add($p)
                    $res.Bytes += $len
                }
            }
            return $res
        } catch {
            $res.Files.Clear()
            $res.Bytes = [int64]0
            $res.Scanned = 0
            $res.TotalBytes = [int64]0
        }
    }

    # 回退路径：逐层遍历，单个子目录失败不影响整棵树
    $stack = New-Object 'System.Collections.Generic.Stack[string]'
    $stack.Push($rootExt)
    while ($stack.Count -gt 0) {
        $cur = $stack.Pop()
        try {
            foreach ($p0 in [IO.Directory]::GetFiles($cur)) {
                $len = [int64](New-Object IO.FileInfo $p0).Length
                $res.Scanned++
                $res.TotalBytes += $len
                if ($len -lt $LimitBytes) {
                    $p = $p0
                    if ($p.StartsWith('\\?\')) { $p = $p.Substring(4) }
                    [void]$res.Files.Add($p)
                    $res.Bytes += $len
                }
            }
            foreach ($sd in [IO.Directory]::GetDirectories($cur)) { $stack.Push($sd) }
        } catch {
            $res.SkippedDirs++
        }
    }
    return $res
}

# 以管理员身份重新启动本脚本
function Start-ElevatedSelf {
    param([int]$LimitKB, [string[]]$Targets)
    $argList = @(
        '-NoProfile', '-ExecutionPolicy', 'Bypass',
        '-File', ('"{0}"' -f $ScriptPath),
        '-AllUsers', '-Elevated', '-MaxSizeKB', "$LimitKB", '-Lang', $LangCode
    )
    # -File 模式下同名参数不能重复出现，多个目录用分号拼成一个参数
    if ($Targets -and $Targets.Count -gt 0) {
        $argList += '-TargetPath'
        $argList += ('"{0}"' -f ($Targets -join ';'))
    }
    try {
        Start-Process -FilePath 'powershell.exe' -Verb RunAs -ArgumentList $argList | Out-Null
        return $true
    } catch {
        return $false
    }
}

# ---------------------------------------------------------------- 主流程

function Invoke-Main {
    $scriptDir = $ScriptDir
    if (-not $scriptDir) { $scriptDir = (Get-Location).Path }

    # 目标目录来源：命令行参数 > 配置块。参数支持数组，或用分号分隔多个目录
    $rawTargets = @()
    if ($TargetPath -and $TargetPath.Count -gt 0) {
        foreach ($item in $TargetPath) {
            if (Test-Blank $item) { continue }
            foreach ($part in ($item -split ';')) { $rawTargets += $part }
        }
    } else {
        $rawTargets = $ConfigTargets
    }
    $limitKB = if ($MaxSizeKB -gt 0) { $MaxSizeKB } else { $ConfigMaxSizeKB }
    if ($limitKB -le 0) { $limitKB = 30 }
    $limitBytes = [int64]$limitKB * 1024

    # 规范化 + 去重 + 按长度倒序（嵌套目录时命中更具体的那个）
    $targets = New-Object 'System.Collections.Generic.List[object]'
    foreach ($raw in $rawTargets) {
        $norm = Convert-ToNormalPath -Path $raw -BaseDir $scriptDir
        if (-not $norm) { continue }
        $exists = $false
        foreach ($t in $targets) { if ($t.Path -ieq $norm) { $exists = $true; break } }
        if ($exists) { continue }
        [void]$targets.Add((New-Object PSObject -Property @{ Path = $norm }))
    }
    if ($targets.Count -eq 0) {
        Write-Host $UI.NoTargets -ForegroundColor Red
        return 1
    }
    $targets = @($targets | Sort-Object { $_.Path.Length } -Descending)

    Write-Host ''
    Write-Host '============================================' -ForegroundColor Cyan
    Write-Host $UI.Title -ForegroundColor Cyan
    Write-Host '============================================' -ForegroundColor Cyan
    Write-Host $UI.Targets
    foreach ($t in $targets) {
        $mark = if ([IO.Directory]::Exists($t.Path)) { '' } else { $UI.TargetMissing }
        Write-Host ('  - ' + $t.Path + $mark)
    }
    Write-Host ($UI.SizeCondition -f $limitKB)
    Write-Host ''

    $isAdmin = Test-IsAdmin
    if ($AllUsers -and -not $isAdmin -and -not $Elevated) {
        Write-Host $UI.RequestingAdmin -ForegroundColor Yellow
        $ok = Start-ElevatedSelf -LimitKB $limitKB -Targets @($targets | ForEach-Object { $_.Path })
        if ($ok) {
            Write-Host $UI.ElevatedStarted -ForegroundColor Green
        } else {
            Write-Host $UI.ElevateFailed -ForegroundColor Red
        }
        return 0
    }

    # ---------------- 扫描回收站 ----------------
    Write-Host $UI.Scanning -ForegroundColor DarkGray
    $swScan = [Diagnostics.Stopwatch]::StartNew()

    $entries = New-Object 'System.Collections.Generic.List[object]'
    $sidCount = 0
    $sidDenied = 0
    $indexTotal = 0
    $indexBad = 0

    foreach ($drive in [IO.DriveInfo]::GetDrives()) {
        if (-not $drive.IsReady) { continue }
        $dt = $drive.DriveType
        if ($dt -ne [IO.DriveType]::Fixed -and $dt -ne [IO.DriveType]::Removable -and $dt -ne [IO.DriveType]::Ram) { continue }

        $recycleRoot = Join-Path $drive.RootDirectory.FullName '$Recycle.Bin'
        if (-not [IO.Directory]::Exists($recycleRoot)) { continue }

        $sidDirs = $null
        try { $sidDirs = [IO.Directory]::GetDirectories($recycleRoot) } catch { continue }

        foreach ($sidDir in $sidDirs) {
            $sidCount++
            $files = $null
            try {
                $files = [IO.Directory]::GetFiles($sidDir)
            } catch {
                $sidDenied++
                continue
            }

            # 一次性取出该 SID 下的子目录名，用于判断条目是文件还是文件夹
            $dirSet = New-Object 'System.Collections.Generic.HashSet[string]' -ArgumentList ([StringComparer]::OrdinalIgnoreCase)
            try {
                foreach ($d in [IO.Directory]::GetDirectories($sidDir)) { [void]$dirSet.Add($d) }
            } catch { }

            foreach ($indexFile in $files) {
                $name = [IO.Path]::GetFileName($indexFile)
                if ($name.Length -lt 3 -or $name[0] -ne [char]'$' -or
                    ($name[1] -ne [char]'I' -and $name[1] -ne [char]'i')) { continue }

                $indexTotal++
                $meta = $null
                try { $meta = Read-IndexEntry -IndexPath $indexFile } catch { $indexBad++; continue }
                if (-not $meta) { $indexBad++; continue }

                # 归属到最长匹配的目标目录；不匹配则完全跳过
                $hitTarget = $null
                foreach ($t in $targets) {
                    if (Test-PathUnder -Path $meta.OriginalPath -Prefix $t.Path) { $hitTarget = $t.Path; break }
                }
                if (-not $hitTarget) { continue }

                $dataPath = Join-Path $sidDir ('$R' + $name.Substring(2))
                [void]$entries.Add((New-Object PSObject -Property @{
                    Target       = $hitTarget
                    SidDir       = $sidDir
                    IndexPath    = $indexFile
                    DataPath     = $dataPath
                    OriginalPath = $meta.OriginalPath
                    Size         = $meta.Size
                    DeletedAt    = $meta.DeletedAt
                    IsDirectory  = $dirSet.Contains($dataPath)
                }))
            }
        }
    }
    $swScan.Stop()

    # ---------------- 统计 ----------------
    # 目录条目不能把“整个目录大小”当成一个文件；递归进去按文件逐个判断大小。
    Write-Host $UI.CountingNested -ForegroundColor DarkGray
    $swNested = [Diagnostics.Stopwatch]::StartNew()
    $nestedScanned = 0
    $nestedSkipped = 0
    foreach ($e in $entries) {
        if (-not $e.IsDirectory) { continue }
        $nested = Get-SmallFilesUnder -Root $e.DataPath -LimitBytes $limitBytes
        $e | Add-Member -MemberType NoteProperty -Name NestedFiles      -Value $nested.Files      -Force
        $e | Add-Member -MemberType NoteProperty -Name NestedBytes      -Value $nested.Bytes      -Force
        $e | Add-Member -MemberType NoteProperty -Name NestedCount      -Value $nested.Files.Count -Force
        $e | Add-Member -MemberType NoteProperty -Name NestedScanned    -Value $nested.Scanned    -Force
        $e | Add-Member -MemberType NoteProperty -Name NestedTotalBytes -Value $nested.TotalBytes -Force
        $nestedScanned += $nested.Scanned
        $nestedSkipped += $nested.SkippedDirs
    }
    $swNested.Stop()

    $stats = @{}
    foreach ($t in $targets) {
        $stats[$t.Path] = New-Object PSObject -Property @{
            LooseFiles = 0; LooseBytes = [int64]0
            LooseSkip = 0;  LooseSkipBytes = [int64]0
            Dirs = 0; DirBytes = [int64]0
            NestedFiles = 0; NestedBytes = [int64]0
            NestedSkip = 0;  NestedSkipBytes = [int64]0
        }
    }
    foreach ($e in $entries) {
        $s = $stats[$e.Target]
        if ($e.IsDirectory) {
            $s.Dirs++
            $s.DirBytes += $e.Size
            $s.NestedFiles += $e.NestedCount
            $s.NestedBytes += $e.NestedBytes
            $s.NestedSkip += ($e.NestedScanned - $e.NestedCount)
            $s.NestedSkipBytes += ($e.NestedTotalBytes - $e.NestedBytes)
        } elseif ($e.Size -lt $limitBytes) {
            $s.LooseFiles++
            $s.LooseBytes += $e.Size
        } else {
            $s.LooseSkip++
            $s.LooseSkipBytes += $e.Size
        }
    }

    $totalLoose = 0; $totalLooseBytes = [int64]0
    $totalNested = 0; $totalNestedBytes = [int64]0
    $totalSkip = 0; $totalSkipBytes = [int64]0
    $totalDirs = 0; $totalDirBytes = [int64]0
    foreach ($t in $targets) {
        $s = $stats[$t.Path]
        $totalLoose += $s.LooseFiles;   $totalLooseBytes += $s.LooseBytes
        $totalNested += $s.NestedFiles; $totalNestedBytes += $s.NestedBytes
        $totalSkip += ($s.LooseSkip + $s.NestedSkip)
        $totalSkipBytes += ($s.LooseSkipBytes + $s.NestedSkipBytes)
        $totalDirs += $s.Dirs;          $totalDirBytes += $s.DirBytes
    }
    $totalFiles = $totalLoose + $totalNested
    $totalFileBytes = $totalLooseBytes + $totalNestedBytes

    Write-Host ($UI.ScanDone -f `
        $sidCount, $indexTotal, $indexBad, $swScan.Elapsed.TotalSeconds) -ForegroundColor DarkGray
    if ($totalDirs -gt 0) {
        Write-Host ($UI.NestedDone -f `
            $totalDirs, $nestedScanned, $swNested.Elapsed.TotalSeconds,
            $(if ($nestedSkipped -gt 0) { $UI.NestedDoneSkip -f $nestedSkipped } else { '' })) -ForegroundColor DarkGray
    }

    if ($sidDenied -gt 0 -and -not $isAdmin) {
        Write-Host ($UI.DeniedHint1 -f $sidDenied) -ForegroundColor DarkYellow
        Write-Host $UI.DeniedHint2 -ForegroundColor DarkYellow
    }
    Write-Host ''

    $idx = 0
    foreach ($t in $targets) {
        $idx++
        $s = $stats[$t.Path]
        $fNum = $s.LooseFiles + $s.NestedFiles
        $fBytes = $s.LooseBytes + $s.NestedBytes
        $kNum = $s.LooseSkip + $s.NestedSkip
        $kBytes = $s.LooseSkipBytes + $s.NestedSkipBytes
        Write-Host ($UI.PerTarget -f $idx, $t.Path) -ForegroundColor White
        Write-Host ($UI.FilesLine -f `
            $fNum, (Format-Size $fBytes), $kNum, (Format-Size $kBytes), $limitKB)
        Write-Host ($UI.DirsLine -f $s.Dirs, (Format-Size $s.DirBytes))
    }
    Write-Host ('--------------------------------------------')
    Write-Host ($UI.TotalFilesLine -f `
        $totalFiles, (Format-Size $totalFileBytes), $totalSkip, (Format-Size $totalSkipBytes)) -ForegroundColor Cyan
    Write-Host ($UI.TotalDirsLine -f $totalDirs, (Format-Size $totalDirBytes)) -ForegroundColor Cyan
    Write-Host ''

    if ($totalFiles -eq 0 -and $totalDirs -eq 0) {
        Write-Host $UI.NothingFound -ForegroundColor Green
        return 0
    }

    # ---------------- 交互选择 ----------------
    $looseFiles = @($entries | Where-Object { -not $_.IsDirectory -and $_.Size -lt $limitBytes })
    $dirEntries = @($entries | Where-Object { $_.IsDirectory })

    Write-Host $UI.MenuTitle
    Write-Host ($UI.Menu1 -f $totalFiles, (Format-Size $totalFileBytes))
    Write-Host ($UI.Menu2 -f `
        $totalFiles, (Format-Size $totalFileBytes), $totalDirs, (Format-Size $totalDirBytes))
    Write-Host $UI.Menu3
    $choice = Read-UserInput $UI.MenuPrompt

    if ($choice -ne '1' -and $choice -ne '2') {
        Write-Host $UI.Cancelled -ForegroundColor Cyan
        return 0
    }

    $delFileCount = $totalFiles
    $delDirCount = if ($choice -eq '2') { $totalDirs } else { 0 }
    if ($choice -eq '2') {
        Write-Host ''
        Write-Host ($UI.DirWarning -f $limitKB) -ForegroundColor Red
    }

    Write-Host ''
    Write-Host ($UI.ConfirmSummary -f $delFileCount, $delDirCount) -ForegroundColor Red
    $confirm = Read-UserInput $UI.ConfirmPrompt
    if ($confirm -ne 'y' -and $confirm -ne 'Y') {
        Write-Host $UI.Cancelled -ForegroundColor Cyan
        return 0
    }

    # ---------------- 删除 ----------------
    $swDel = [Diagnostics.Stopwatch]::StartNew()
    $okLoose = 0; $okNested = 0; $okDirs = 0; $nestedCovered = 0
    $failures = New-Object 'System.Collections.Generic.List[string]'

    # 1) 散落文件：$R 实体 + $I 元数据成对删除
    foreach ($e in $looseFiles) {
        try {
            # 安全护栏：只允许操作 $Recycle.Bin 内的 $R / $I
            if (-not $e.DataPath.StartsWith($e.SidDir + '\', [StringComparison]::OrdinalIgnoreCase) -or
                -not $e.IndexPath.StartsWith($e.SidDir + '\', [StringComparison]::OrdinalIgnoreCase)) {
                throw $UI.PathGuardFail
            }

            Remove-RecycleData -Path $e.DataPath
            Remove-FileForce -Path $e.IndexPath
            $okLoose++
        } catch {
            $msg = $_.Exception.Message
            if (Test-Blank $msg) { $msg = $_.Exception.GetType().Name }
            [void]$failures.Add(('{0}  [{1}]' -f $e.OriginalPath, $msg))
        }
    }

    if ($choice -eq '1') {
        # 2) 只删目录条目内部符合条件的小文件，目录条目本身保留
        foreach ($d in $dirEntries) {
            foreach ($f in $d.NestedFiles) {
                try {
                    if (-not $f.StartsWith($d.DataPath + '\', [StringComparison]::OrdinalIgnoreCase)) {
                        throw $UI.PathGuardFail
                    }
                    Remove-FileForce -Path $f
                    $okNested++
                } catch {
                    $msg = $_.Exception.Message
                    if (Test-Blank $msg) { $msg = $_.Exception.GetType().Name }
                    [void]$failures.Add(('{0}  [{1}]' -f $f, $msg))
                }
            }
        }
    } else {
        # 2) 目录条目整体删除，内部文件随之删除
        foreach ($d in $dirEntries) {
            try {
                if (-not $d.DataPath.StartsWith($d.SidDir + '\', [StringComparison]::OrdinalIgnoreCase) -or
                    -not $d.IndexPath.StartsWith($d.SidDir + '\', [StringComparison]::OrdinalIgnoreCase)) {
                    throw $UI.PathGuardFail
                }

                Remove-RecycleData -Path $d.DataPath
                Remove-FileForce -Path $d.IndexPath
                $okDirs++
                $nestedCovered += $d.NestedCount
            } catch {
                $msg = $_.Exception.Message
                if (Test-Blank $msg) { $msg = $_.Exception.GetType().Name }
                [void]$failures.Add(('{0}  [{1}]' -f $d.OriginalPath, $msg))
            }
        }
    }
    $swDel.Stop()

    Write-Host ''
    if ($choice -eq '1') {
        Write-Host ($UI.DoneKeepDirs -f `
            ($okLoose + $okNested), $dirEntries.Count, $failures.Count, $swDel.Elapsed.TotalSeconds) -ForegroundColor Green
    } else {
        Write-Host ($UI.DoneWithDirs -f `
            ($okLoose + $nestedCovered), $okDirs, $failures.Count, $swDel.Elapsed.TotalSeconds) -ForegroundColor Green
    }

    if ($failures.Count -gt 0) {
        Write-Host $UI.FailHeader -ForegroundColor Red
        $show = [Math]::Min(20, $failures.Count)
        for ($i = 0; $i -lt $show; $i++) { Write-Host ('  - ' + $failures[$i]) -ForegroundColor Red }
        if ($failures.Count -gt $show) { Write-Host ($UI.FailMore -f ($failures.Count - $show)) -ForegroundColor Red }
        if (-not $isAdmin) {
            Write-Host $UI.FailHintAdmin -ForegroundColor DarkYellow
        } else {
            Write-Host $UI.FailHintBusy -ForegroundColor DarkYellow
        }
    }
    return 0
}

$exitCode = 1
try {
    $exitCode = Invoke-Main
} catch {
    Write-Host ''
    Write-Host ($UI.UnhandledError + $_.Exception.Message) -ForegroundColor Red
    Write-Host ($UI.ErrorAt + $_.InvocationInfo.PositionMessage) -ForegroundColor DarkGray
    $exitCode = 1
} finally {
    if (-not $NoPause) {
        Write-Host ''
        [void](Read-UserInput $UI.PressEnter)
    }
}
exit $exitCode
