English | [简体中文](README.zh-CN.md)

# trash-path-cleaner

A small Windows tool that permanently deletes recycle-bin items by their **original path**: only files that were deleted from the folders you specify (or their subfolders) and whose original size is below a threshold are removed. Everything else is left untouched.

It reads the raw `$I` metadata and `$R` payload of the recycle bin instead of enumerating entries through Shell COM, so thousands of records are scanned in well under a second.

- Chinese name: 回收站定点清理
- English name: trash-path-cleaner
- License: MIT

---

## Features

- **Path-scoped cleanup**: only entries under the configured folders are touched; `C:\a` never matches `C:\ab` because prefix matching respects the path separator.
- **Size filter**: a file is deleted only when its original size is strictly smaller than the threshold.
- **Multiple folders**: several target folders can be configured; each entry is attributed to the longest matching folder and reported separately.
- **Files vs. folders**: when a whole folder was deleted, the tool recursively enumerates the files inside its `$R` directory and judges each file by its own size instead of treating the folder size as one file.
- **Two delete modes**: small files only (folder structure kept) or files plus folder entries.
- **Fast**: a single directory pass reads every `$I`, parsing only the header and path — no COM calls per item.
- **Bilingual UI**: follows the system UI language automatically; override with `-Lang zh|en`.
- **Legacy friendly**: works from PowerShell 2.0; supports the Win7 `$I` V1 (544-byte) and Win10 1809+ V2 formats.
- **External config**: target folders and threshold live in `config.env`, so the script never needs editing. The file is not shipped with the repository (it is listed in `.gitignore`); when it is missing the script exits with the required keys and a sample.
- **Safety guard**: only `$I` / `$R` items inside `$Recycle.Bin\<SID>\` can be deleted; files at their original locations are never touched.

## Requirements

| Item | Requirement |
| --- | --- |
| OS | Windows 11 / Windows 10 / Windows 7 (Vista should work but is untested) |
| PowerShell | 2.0 or later (Win7 ships 2.0, Win10/11 ship 5.1) |
| .NET Framework | 3.5 or later (bundled with the OS, nothing to install) |
| Privileges | Normal user for your own recycle bin; administrator to process other users' or system accounts |

Notes: `EnumerateFiles` requires .NET 4.0+, so PowerShell 2.0 on Win7 automatically falls back to a `GetFiles` path list — same behavior, slightly slower. Subfolders that cannot be read are skipped instead of failing the whole tree.

The configuration file is plain `key=value` text, not JSON, so it does not depend on `ConvertFrom-Json` (PowerShell 3.0+).

## Install / Download

1. Download or clone this repository.
2. Keep `trash-path-cleaner.ps1` and `run.bat` in the same folder (the launcher looks for the script by a fixed name).
3. Create `config.env` next to the script using the sample in section 2 (the repository does not include it and Git ignores it), then double-click `run.bat`.
4. If Windows marked files downloaded from the internet, unblock them in file properties, or right-click `run.bat` and choose "Run as administrator".

```
trash-path-cleaner/
├─ trash-path-cleaner.ps1   # main script
├─ run.bat                  # double-click launcher
├─ config.env               # create it yourself (ignored by Git)
├─ README.md                # English documentation (this file)
├─ README.zh-CN.md          # Chinese documentation
├─ LICENSE
└─ .gitignore
```

## Usage

### 1. Double-click (recommended)

Double-click `run.bat`. The script runs in the current console window and waits for Enter before closing.

### 2. Create / edit config.env

Create a plain text file named `config.env` next to the script (not included in the repository, ignored by Git):

```ini
# Target folders, separate multiple paths with ';'
TARGETS=C:\Users\username\test_directory

# Size threshold in KB: only files strictly smaller than this are deleted
MAX_SIZE_KB=30
```

- `TARGETS`: replace with the folders you want to clean; separate several with `;`. Environment variables such as `%USERPROFILE%\Downloads` and relative paths (relative to the script folder) are supported. The folder does not need to exist — historical recycle-bin entries are still matched by path prefix.
- `MAX_SIZE_KB`: positive integer, in KB.
- Lines starting with `#` or `;` are comments; quotes are not required even when a path contains spaces.
- If `config.env` is missing at startup, the script prints the required keys and a sample, then exits — it never creates the file for you.
- `config.env` is listed in `.gitignore`, so a fresh clone starts without it; copy the sample above into your own file.

### 3. Command-line parameters (override the config file)

| Parameter | Description |
| --- | --- |
| `-TargetPath <path>` | Target folders, separated by `;`, e.g. `'D:\proj;E:\tmp'`; overrides `TARGETS` |
| `-MaxSizeKB <int>` | Size threshold in KB; overrides `MAX_SIZE_KB` |
| `-AllUsers` | Requests administrator rights and checks every user's recycle bin |
| `-Lang zh\|en` | Forces the Chinese or English UI; otherwise follows the system UI language |
| `-NoPause` | Does not wait for Enter at the end (handy for automation) |

Examples:

```powershell
.\trash-path-cleaner.ps1
.\trash-path-cleaner.ps1 -TargetPath 'D:\proj' -MaxSizeKB 100
.\trash-path-cleaner.ps1 -TargetPath 'D:\proj;E:\tmp'
.\trash-path-cleaner.ps1 -AllUsers
.\trash-path-cleaner.ps1 -Lang en
```

Parameters can also be passed through the launcher: `run.bat -TargetPath "D:\proj" -MaxSizeKB 100 -Lang en`.

### 4. Reading the report

The script scans first and asks questions later; nothing is deleted during the scan. Each target folder gets two lines:

```
[1] C:\Users\username\test_directory
     Files: to delete 23 (2.1 KB); to keep 169 (1.07 GB) (size >= 30 KB)
     Folders: 0 (0 B) (deleted only by option 2; option 1 deletes small files inside but keeps the folder)
```

- **Files**: `to delete` counts files whose path matches and whose size is below the threshold (including nested files found inside deleted folders); `to keep` counts matching files that are at or above the threshold.
- **Folders**: the number of matched "whole folder deleted" entries and their original total size. Option 1 keeps the folder and deletes only qualifying files inside; option 2 deletes the whole entry.

### 5. The three options

```
Choose an action:
  1  Delete files only: 33 (125.8 KB) (folders kept)
  2  Delete files and folders: files 33 / 125.8 KB, folders 3 / 0 B
  3  Cancel (no changes)
```

- **Option 1**: files only. Standalone files lose their `$R` + `$I` pair; inside folder entries, only recursively detected small files are removed and the folder itself is kept (empty folders may remain).
- **Option 2**: files and folder entries together. **Everything** inside those folders (including files above the threshold) is permanently deleted with the folder.
- **Option 3**: cancel, no changes.

After choosing 1 or 2 you must type `y` to confirm; anything else cancels.

## Permissions

- Recycle bins are stored per SID: `X:\$Recycle.Bin\<SID>\`. With normal rights only your own SID folder is readable; folders of other users and system accounts are skipped and their count is reported.
- To include them, add `-AllUsers` (triggers UAC elevation) or right-click `run.bat` and choose "Run as administrator".
- Entries that still fail after elevation are usually locked by a running program; close it and retry.

## Performance

Measured on Windows 11 with PowerShell 5.1 and 4 SID recycle bins:

| Stage | Volume | Time |
| --- | --- | --- |
| Scan `$I` | 312 metadata entries (6 unreadable) | 0.38 – 0.43 s |
| Recursive count inside folder entries | 3 folder entries / 0 files | 0.04 – 0.05 s |

- Scanning is O(entries); each entry is read once and only the header plus path is parsed — no COM calls.
- Folder recursion enumerates only once: `EnumerateFiles` on .NET 4.0+, falling back to a per-directory walk only when a subfolder is inaccessible.
- The config file is read once at startup, and command-line overrides are resolved at startup too, so scanning performance is unaffected.
- Win7 + PowerShell 2.0 uses a `GetFiles` path list, which is slower with many entries but behaves identically.

## Known limitations

- Only the `$Recycle.Bin` layout is supported (NTFS and most NTFS-formatted removable drives). The legacy FAT32/exFAT `RECYCLER` + `INFO2` layout is not parsed and its entries are ignored.
- Unrecognized `$I` files (corrupted or not V1/V2) are counted as "unreadable" and skipped; they are never deleted.
- Deletion is permanent, does not go to the recycle bin, and cannot be undone.
- Option 1 may leave empty folders after removing small files; use option 2 if you want the folder structure gone too.
- The size comparison is strictly "smaller than"; files exactly at the threshold are kept.
- Only local fixed/removable drives are scanned; network drives are ignored.
- `-AllUsers` relies on UAC; if elevation is cancelled the script continues with the current rights.
- The config file does not support sections (`[section]`), nesting, or arrays; only the `TARGETS` and `MAX_SIZE_KB` keys are read, and unknown keys are ignored.
