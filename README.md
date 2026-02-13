# StoragePoolTray

A lightweight **Windows 11** tray app (pure **PowerShell 5.1 + WPF**) that shows **Storage Spaces / Storage Pool health** in a draggable desktop widget.

## Features

- Tray icon with menu: **Show/Hide**, **Pin/Unpin**, **Refresh Now**, **Exit**
- Widget shows **on launch**
- Auto-resizes to content (with MaxWidth/MaxHeight; ScrollViewer handles overflow)
- **Collapse/Expand each pool** (state persisted)
- **Pinned** = always on top (Topmost)
- **Unpinned** = pushed behind other windows (background widget)
- Health-based coloring
- Disk type labels: `[NVMe] [SSD] [HDD] [VD]` (ASCII-safe)
- Unpooled disks section
- Footer: **Last updated** timestamp
- Saves position/settings to: `%APPDATA%\StoragePoolTray\StoragePoolWidget.json`

## Requirements

- Windows 11
- Windows PowerShell 5.1 (STA)
- Storage cmdlets available (`Get-PhysicalDisk`, `Get-StoragePool`, etc.)
- CIM access to `root\Microsoft\Windows\Storage`

## Run

From Windows PowerShell (not admin):

```powershell
Set-Location "$env:USERPROFILE\Documents\StoragePoolTray"
powershell.exe -Sta -NoProfile -ExecutionPolicy Bypass -File .\StoragePoolTray.ps1
```

## Auto-start on login

### Option A: Hidden startup (recommended)

1. Double-click `RunHidden.vbs` once to test.
2. Press `Win + R`, type `shell:startup`
3. Create a shortcut to:

```text
wscript.exe "C:\Users\alex\Documents\StoragePoolTray\RunHidden.vbs"
```

### Option B: Visible PowerShell window

Create a shortcut with target:

```text
C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe -Sta -NoProfile -ExecutionPolicy Bypass -File "C:\Users\alex\Documents\StoragePoolTray\StoragePoolTray.ps1"
```

And set **Start in** to:

```text
C:\Users\alex\Documents\StoragePoolTray
```

## Troubleshooting

### Startup runs but nothing shows
- Unblock the script:

```powershell
Unblock-File .\StoragePoolTray.ps1
```

### “Blank window pops” when minimized
Fixed by avoiding `SWP_SHOWWINDOW` and skipping z-order operations while minimized.

### Desktop freeze while dragging
Fixed by using a non-blocking dispatcher pump (`BeginInvoke`) and debounced config saves.

## License

MIT — see `LICENSE`.
