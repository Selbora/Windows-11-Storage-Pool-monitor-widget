#requires -version 5.1
<#
StoragePoolTray.ps1 (PowerShell 5.1)
Tray app + Storage Spaces widget (WPF)

Features:
- Shows on launch
- Auto-resizes to content (with MaxWidth/MaxHeight; ScrollViewer handles overflow)
- Expand/Collapse per pool (persisted)
- Pin/Unpin:
    Pinned   = Topmost (no push-to-background)
    Unpinned = Not topmost + pushed behind other windows
- Remembers position + settings in %APPDATA%\StoragePoolTray\StoragePoolWidget.json
- Refreshes every minute
- Disk “icons” ASCII-safe: [NVMe] [SSD] [HDD] [VD]
- Clean exit (no phantom windows), saves config on exit, no flashing
- No blank window popping on refresh/minimize
- No desktop freeze on drag (BeginInvoke + debounced saves)
- Footer shows "Last updated" timestamp
#>

# -----------------------------
# Ensure STA (WPF requirement)
# -----------------------------
if ([System.Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
    Start-Process powershell -ArgumentList @(
        '-Sta','-NoProfile','-ExecutionPolicy','Bypass','-File',"`"$PSCommandPath`""
    ) | Out-Null
    exit
}

# -----------------------------
# Assemblies
# -----------------------------
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# -----------------------------
# Win32: push window to background
# NOTE: We DO NOT use SWP_SHOWWINDOW to avoid popping blank windows.
# -----------------------------
Add-Type @"
using System;
using System.Runtime.InteropServices;
public static class Win32 {
  [DllImport("user32.dll")]
  public static extern bool SetWindowPos(
    IntPtr hWnd, IntPtr hWndInsertAfter,
    int X, int Y, int cx, int cy, uint flags);

  public static readonly IntPtr HWND_BOTTOM = new IntPtr(1);
  public const uint SWP_NOMOVE = 0x0002;
  public const uint SWP_NOSIZE = 0x0001;
  public const uint SWP_NOACTIVATE = 0x0010;
}
"@

# -----------------------------
# App folder + config
# -----------------------------
$AppDir     = Join-Path $env:APPDATA 'StoragePoolTray'
$null = New-Item -ItemType Directory -Path $AppDir -Force -ErrorAction SilentlyContinue
$ConfigPath = Join-Path $AppDir 'StoragePoolWidget.json'

$Config = @{
    Left = $null
    Top  = $null
    Pinned = $ indicate = $false
    PoolExpanded = @{}   # key: pool.ObjectId -> bool
}

function Load-Config {
    if (Test-Path $ConfigPath) {
        try {
            $c = Get-Content $ConfigPath -Raw | ConvertFrom-Json
            if ($null -ne $c.Left)   { $Config.Left = [double]$c.Left }
            if ($null -ne $c.Top)    { $Config.Top  = [double]$c.Top }
            if ($null -ne $c.Pinned) { $Config.Pinned = [bool]$c.Pinned }

            if ($null -ne $c.PoolExpanded) {
                $Config.PoolExpanded = @{}
                $c.PoolExpanded.PSObject.Properties | ForEach-Object {
                    $Config.PoolExpanded[$_.Name] = [bool]$_.Value
                }
            }
        } catch {}
    }
}

function Save-Config {
    try {
        @{
            Left = $Config.Left
            Top  = $Config.Top
            Pinned = $Config.Pinned
            PoolExpanded = $Config.PoolExpanded
        } | ConvertTo-Json -Depth 10 | Set-Content -Encoding UTF8 $ConfigPath
    } catch {}
}

Load-Config

# Debounced save timer (prevents disk write spam during dragging)
$script:SaveTimer = New-Object System.Windows.Threading.DispatcherTimer
$script:SaveTimer.Interval = [TimeSpan]::FromMilliseconds(500)
$script:SaveTimer.Add_Tick({
    $script:SaveTimer.Stop()
    Save-Config
})

# -----------------------------
# Health normalization + Brushes
# -----------------------------
function Normalize-HealthStatus {
    param($hs)

    if ($null -eq $hs) { return "Unknown" }

    if ($hs -is [System.Array]) {
        $hs = ($hs | Select-Object -First 1)
        if ($null -eq $hs) { return "Unknown" }
    }

    if ($hs -is [ValueType]) {
        switch ([int]$hs) {
            1 { return "Healthy" }
            2 { return "Warning" }
            3 { return "Unhealthy" }
            default { return "Unknown" }
        }
    }

    $s = ([string]$hs).Trim()
    $n = 0
    if ([int]::TryParse($s, [ref]$n)) {
        switch ($n) {
            1 { return "Healthy" }
            2 { return "Warning" }
            3 { return "Unhealthy" }
            default { return "Unknown" }
        }
    }

    switch -Regex ($s) {
        '^Healthy$'   { return "Healthy" }
        '^OK$'        { return "Healthy" }
        '^Warning$'   { return "Warning" }
        '^Degraded$'  { return "Warning" }
        '^Unhealthy$' { return "Unhealthy" }
        '^Failed$'    { return "Unhealthy" }
        default       { return $s }
    }
}

function Get-Brush {
    param($hs)
    switch (Normalize-HealthStatus $hs) {
        "Healthy"   { return [System.Windows.Media.Brushes]::LimeGreen }
        "Warning"   { return [System.Windows.Media.Brushes]::Orange }
        "Unhealthy" { return [System.Windows.Media.Brushes]::Red }
        default     { return [System.Windows.Media.Brushes]::LightGray }
    }
}

function Get-NonEmptyName {
    param($obj)
    if ($obj -and $obj.FriendlyName -and ($obj.FriendlyName.ToString().Trim().Length -gt 0)) { return $obj.FriendlyName }
    if ($obj -and $obj.Model -and ($obj.Model.ToString().Trim().Length -gt 0)) { return $obj.Model }
    if ($obj -and $obj.Name -and ($obj.Name.ToString().Trim().Length -gt 0)) { return $obj.Name }
    if ($obj -and $obj.SerialNumber -and ($obj.SerialNumber.ToString().Trim().Length -gt 0)) { return $obj.SerialNumber }
    return "(unknown)"
}

# -----------------------------
# Disk icon logic (ASCII-safe)
# -----------------------------
function Build-PhysicalDiskLookup {
    $byUniqueId = @{}
    $bySerial   = @{}
    $byName     = @{}

    $cmdDisks = @(Get-PhysicalDisk -ErrorAction SilentlyContinue)
    foreach ($d in $cmdDisks) {
        $u = [string]$d.UniqueId
        $s = [string]$d.SerialNumber
        $n = [string]$d.FriendlyName

        if ($u -and -not $byUniqueId.ContainsKey($u)) { $byUniqueId[$u] = $d }
        if ($s -and -not $bySerial.ContainsKey($s))   { $bySerial[$s]   = $d }
        if ($n -and -not $byName.ContainsKey($n))     { $byName[$n]     = $d }
    }

    return @{
        ByUniqueId = $byUniqueId
        BySerial   = $bySerial
        ByName     = $byName
    }
}

function Get-PhysicalDiskIcon {
    param($cimDisk, $lookup)

    $bus  = ""
    $media = ""

    $u = [string]$cimDisk.UniqueId
    $s = [string]$cimDisk.SerialNumber
    $n = [string]$cimDisk.FriendlyName

    $match = $null
    if ($u -and $lookup.ByUniqueId.ContainsKey($u)) { $match = $lookup.ByUniqueId[$u] }
    elseif ($s -and $lookup.BySerial.ContainsKey($s)) { $match = $lookup.BySerial[$s] }
    elseif ($n -and $lookup.ByName.ContainsKey($n)) { $match = $lookup.ByName[$n] }

    if ($match) {
        $bus   = ([string]$match.BusType).Trim()
        $media = ([string]$match.MediaType).Trim()
    } else {
        $bus   = ([string]$cimDisk.BusType).Trim()
        $media = ([string]$cimDisk.MediaType).Trim()
    }

    $busU   = $bus.ToUpperInvariant()
    $mediaU = $media.ToUpperInvariant()

    if ($busU -match 'NVME') { return "[NVMe]" }
    if ($mediaU -match 'SSD') { return "[SSD]" }
    if ($mediaU -match 'HDD') { return "[HDD]" }
    if ($busU -match 'SATA|ATA|SAS|USB|SD') { return "[HDD]" }

    return "[?]"
}

function Get-VirtualDiskIcon { return "[VD]" }

# -----------------------------
# Storage via CIM (pooled) + CIM unpooled
# -----------------------------
$NS = 'root/Microsoft/Windows/Storage'

function Get-PoolsCim {
    @(Get-CimInstance -Namespace $NS -ClassName MSFT_StoragePool -ErrorAction SilentlyContinue |
      Where-Object { $_.IsPrimordial -eq $false })
}

function Get-VirtualDisksForPoolCim {
    param($poolCim)
    @(Get-CimAssociatedInstance -InputObject $poolCim -Namespace $NS -ResultClassName MSFT_VirtualDisk -ErrorAction SilentlyContinue)
}

function Get-PhysicalDisksForPoolCim {
    param($poolCim)
    @(Get-CimAssociatedInstance -InputObject $poolCim -Namespace $NS -ResultClassName MSFT_PhysicalDisk -ErrorAction SilentlyContinue)
}

function Get-UnpooledPhysicalDisks {
    $pools = Get-PoolsCim

    $pooledIds = New-Object 'System.Collections.Generic.HashSet[string]'
    foreach ($p in $pools) {
        $pds = @(Get-PhysicalDisksForPoolCim $p)
        foreach ($pd in $pds) {
            if ($pd.ObjectId) { [void]$pooledIds.Add([string]$pd.ObjectId) }
        }
    }

    $all = @(Get-CimInstance -Namespace $NS -ClassName MSFT_PhysicalDisk -ErrorAction SilentlyContinue)
    @($all | Where-Object { -not ($_.ObjectId -and $pooledIds.Contains([string]$_.ObjectId)) })
}

# -----------------------------
# WPF UI helpers (Grid with header)
# -----------------------------
function New-Header($text) {
    $t = New-Object System.Windows.Controls.TextBlock
    $t.Text = $text
    $t.FontSize = 14
    $t.FontWeight = "Bold"
    $t.Margin = "0,10,0,4"
    $t.Foreground = [System.Windows.Media.Brushes]::White
    $t
}

function New-Legend {
    $t = New-Object System.Windows.Controls.TextBlock
    $t.Text = "Legend: [NVMe] NVMe   [SSD] SSD   [HDD] HDD   [?] Unknown   [VD] Virtual"
    $t.FontFamily = "Consolas"
    $t.Foreground = [System.Windows.Media.Brushes]::LightGray
    $t.Margin = "0,0,0,10"
    $t
}

function New-DiskGrid {
    $g = New-Object System.Windows.Controls.Grid
    $g.Margin = "0,2,0,8"
    # Columns: Icon | Name | Health | Extra
    $g.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition -Property @{ Width="Auto" }))
    $g.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition -Property @{ Width="2*" }))
    $g.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition -Property @{ Width="Auto" }))
    $g.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition -Property @{ Width="Auto" }))
    $g
}

function Add-DiskGridHeader {
    param(
        [Parameter(Mandatory=$true)] $Grid,
        [string]$ColName = "Name",
        [string]$ColHealth = "Health",
        [string]$ColExtra = "Extra"
    )

    $Grid.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition)) | Out-Null
    $r = $Grid.RowDefinitions.Count - 1
    $headers = @("", $ColName, $ColHealth, $ColExtra)

    for ($i = 0; $i -lt 4; $i++) {
        $t = New-Object System.Windows.Controls.TextBlock
        $t.Text = $headers[$i]
        $t.FontWeight = "Bold"
        $t.FontFamily = "Consolas"
        $t.Foreground = [System.Windows.Media.Brushes]::LightGray
        $t.Margin = "0,0,14,4"
        [System.Windows.Controls.Grid]::SetRow($t, $r)
        [System.Windows.Controls.Grid]::SetColumn($t, $i)
        $Grid.Children.Add($t) | Out-Null
    }
}

function Add-DiskGridRow {
    param(
        [Parameter(Mandatory=$true)] $Grid,
        [string]$Icon,
        [string]$Name,
        $Health,
        [string]$Extra
    )

    $r = $Grid.RowDefinitions.Count
    $Grid.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition)) | Out-Null

    $hsNorm = Normalize-HealthStatus $Health
    $brush  = Get-Brush $hsNorm

    $vals = @($Icon, $Name, $hsNorm, $Extra)
    for ($i=0; $i -lt 4; $i++) {
        $t = New-Object System.Windows.Controls.TextBlock
        $t.Text = $vals[$i]
        $t.FontFamily = "Consolas"
        $t.Foreground = $brush
        $t.Margin = "0,0,14,0"
        [System.Windows.Controls.Grid]::SetRow($t, $r)
        [System.Windows.Controls.Grid]::SetColumn($t, $i)
        $Grid.Children.Add($t) | Out-Null
    }
}

# -----------------------------
# Create widget window
# -----------------------------
$window = New-Object System.Windows.Window
$window.Title = "Storage Pool Health"

# Auto-size to content
$window.SizeToContent = "WidthAndHeight"
$window.MaxWidth  = 900
$window.MaxHeight = [System.Windows.SystemParameters]::WorkArea.Height - 40

# Pin state
$window.Topmost = [bool]$Config.Pinned

# Tray app: keep off taskbar
$window.ShowInTaskbar = $false
$window.ResizeMode = "NoResize"
$window.WindowStyle = "ToolWindow"
$window.Background = "#1F1F1F"
$window.Foreground = "White"

# Restore position
if ($null -ne $Config.Left -and $null -ne $Config.Top) {
    $window.Left = [double]$Config.Left
    $window.Top  = [double]$Config.Top
}

# Allow close flag (prevents phantom windows on exit)
$script:AllowClose = $false

# Drag to move
$window.Add_MouseLeftButtonDown({ $window.DragMove() })

# Save position (debounced)
$window.Add_LocationChanged({
    $Config.Left = $window.Left
    $Config.Top  = $window.Top
    $script:SaveTimer.Stop()
    $script:SaveTimer.Start()
})

# Closing the window hides it (Exit will set $script:AllowClose = $true)
$window.Add_Closing({
    if (-not $script:AllowClose) {
        $_.Cancel = $true
        $window.Hide()
    }
})

# Content
$scroll = New-Object System.Windows.Controls.ScrollViewer
$scroll.VerticalScrollBarVisibility = "Auto"
$scroll.Margin = "10"
$root = New-Object System.Windows.Controls.StackPanel
$scroll.Content = $root
$window.Content = $scroll

# Footer: last updated
$footerText = New-Object System.Windows.Controls.TextBlock
$footerText.FontFamily = "Consolas"
$footerText.FontSize = 11
$footerText.Foreground = [System.Windows.Media.Brushes]::Gray
$footerText.Margin = "0,12,0,0"
$footerText.HorizontalAlignment = "Right"

# -----------------------------
# Background behavior helpers
# -----------------------------
function Push-To-Background {
    try {
        if (-not $window.IsVisible) { return }
        if ($window.WindowState -eq [System.Windows.WindowState]::Minimized) { return }

        $hwnd = (New-Object System.Windows.Interop.WindowInteropHelper $window).Handle
        [Win32]::SetWindowPos(
            $hwnd,
            [Win32]::HWND_BOTTOM,
            0,0,0,0,
            [Win32]::SWP_NOMOVE -bor
            [Win32]::SWP_NOSIZE -bor
            [Win32]::SWP_NOACTIVATE
        ) | Out-Null
    } catch {}
}

function Maybe-PushToBackground {
    if (-not $Config.Pinned) { Push-To-Background }
}

$window.Add_SourceInitialized({ Maybe-PushToBackground })

# -----------------------------
# Update routine with expand/collapse per pool + footer
# -----------------------------
function Update-StorageInfo {
    if ($window.WindowState -eq [System.Windows.WindowState]::Minimized) { return }

    $root.Children.Clear()
    $root.Children.Add((New-Legend)) | Out-Null

    $lookup = Build-PhysicalDiskLookup
    $pools  = Get-PoolsCim

    if ($pools.Count -eq 0) {
        $root.Children.Add((New-Header "No Storage Pools found")) | Out-Null

        $footerText.Text = "Last updated: {0}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
        $root.Children.Add($footerText) | Out-Null

        Maybe-PushToBackground
        return
    }

    foreach ($pool in $pools) {
        $poolName = if ($pool.FriendlyName) { $pool.FriendlyName } else { $pool.Name }
        $poolId   = [string]$pool.ObjectId

        $poolHealth = Normalize-HealthStatus $pool.HealthStatus
        $hdr = New-Object System.Windows.Controls.TextBlock
        $hdr.Text = ("{0}  -  {1}" -f $poolName, $poolHealth)
        $hdr.FontSize = 14
        $hdr.FontWeight = "Bold"
        $hdr.Foreground = (Get-Brush $poolHealth)

        $exp = New-Object System.Windows.Controls.Expander
        $exp.Header = $hdr
        $exp.Margin = "0,6,0,6"

        $isExpanded = $true
        if ($Config.PoolExpanded.ContainsKey($poolId)) { $isExpanded = [bool]$Config.PoolExpanded[$poolId] }
        $exp.IsExpanded = $isExpanded

        $exp.Add_Expanded({
            $Config.PoolExpanded[$poolId] = $true
            $script:SaveTimer.Stop()
            $script:SaveTimer.Start()
        })
        $exp.Add_Collapsed({
            $Config.PoolExpanded[$poolId] = $false
            $script:SaveTimer.Stop()
            $script:SaveTimer.Start()
        })

        $panel = New-Object System.Windows.Controls.StackPanel
        $panel.Margin = "10,6,0,0"
        $exp.Content = $panel

        # Virtual Disks
        $panel.Children.Add((New-Header "Virtual Disks")) | Out-Null
        $vds = @(Get-VirtualDisksForPoolCim $pool)
        if ($vds.Count -gt 0) {
            $g = New-DiskGrid
            Add-DiskGridHeader -Grid $g -ColName "Name" -ColHealth "Health" -ColExtra "Size"
            foreach ($vd in $vds) {
                $name  = Get-NonEmptyName $vd
                $extra = if ($vd.Size) { "{0:N1} TB" -f ($vd.Size/1TB) } else { "" }
                Add-DiskGridRow -Grid $g -Icon (Get-VirtualDiskIcon) -Name $name -Health $vd.HealthStatus -Extra $extra
            }
            $panel.Children.Add($g) | Out-Null
        } else {
            $panel.Children.Add((New-Header "  (none)")) | Out-Null
        }

        # Physical Disks (Pooled)
        $panel.Children.Add((New-Header "Physical Disks (Pooled)")) | Out-Null
        $pds = @(Get-PhysicalDisksForPoolCim $pool)
        if ($pds.Count -gt 0) {
            $g = New-DiskGrid
            Add-DiskGridHeader -Grid $g -ColName "Name" -ColHealth "Health" -ColExtra "Status"
            foreach ($pd in $pds) {
                $name  = Get-NonEmptyName $pd
                $extra = if ($pd.OperationalStatus) { [string]($pd.OperationalStatus | Select-Object -First 1) } else { "" }
                $icon  = Get-PhysicalDiskIcon -cimDisk $pd -lookup $lookup
                Add-DiskGridRow -Grid $g -Icon $icon -Name $name -Health $pd.HealthStatus -Extra $extra
            }
            $panel.Children.Add($g) | Out-Null
        } else {
            $panel.Children.Add((New-Header "  (none)")) | Out-Null
        }

        $root.Children.Add($exp) | Out-Null
    }

    # Unpooled
    $root.Children.Add((New-Header "Unpooled Physical Disks")) | Out-Null
    $up = @(Get-UnpooledPhysicalDisks)
    if ($up.Count -gt 0) {
        $g = New-DiskGrid
        Add-DiskGridHeader -Grid $g -ColName "Name" -ColHealth "Health" -ColExtra "Status"
        foreach ($pd in $up) {
            $name  = Get-NonEmptyName $pd
            $extra = if ($pd.OperationalStatus) { [string]($pd.OperationalStatus | Select-Object -First 1) } else { "" }
            $icon  = Get-PhysicalDiskIcon -cimDisk $pd -lookup $lookup
            Add-DiskGridRow -Grid $g -Icon $icon -Name $name -Health $pd.HealthStatus -Extra $extra
        }
        $root.Children.Add($g) | Out-Null
    } else {
        $root.Children.Add((New-Header "  (none)")) | Out-Null
    }

    $footerText.Text = "Last updated: {0}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
    $root.Children.Add($footerText) | Out-Null

    Maybe-PushToBackground
}

# -----------------------------
# WPF Refresh Timer
# -----------------------------
$wpfTimer = New-Object System.Windows.Threading.DispatcherTimer
$wpfTimer.Interval = [TimeSpan]::FromMinutes(1)
$wpfTimer.Add_Tick({ Update-StorageInfo | Out-Null })
$wpfTimer.Start()

Update-StorageInfo | Out-Null

# -----------------------------
# Tray Icon + menu
# -----------------------------
$notifyIcon = New-Object System.Windows.Forms.NotifyIcon
$notifyIcon.Text = "Storage Pool Widget"
$notifyIcon.Icon = [System.Drawing.SystemIcons]::Information
$notifyIcon.Visible = $true

$menu = New-Object System.Windows.Forms.ContextMenuStrip

$miShowHide = New-Object System.Windows.Forms.ToolStripMenuItem
$miShowHide.Text = "Hide Widget"
$miShowHide.Add_Click({
    if ($window.IsVisible) {
        $window.Hide()
        $miShowHide.Text = "Show Widget"
    } else {
        $window.Show()
        $window.Activate() | Out-Null
        Maybe-PushToBackground
        $miShowHide.Text = "Hide Widget"
    }
})

$miPin = New-Object System.Windows.Forms.ToolStripMenuItem
$miPin.Text = if ($Config.Pinned) { "Unpin (Allow Background)" } else { "Pin (Always On Top)" }
$miPin.Add_Click({
    $Config.Pinned = -not $Config.Pinned
    $window.Topmost = [bool]$Config.Pinned
    Save-Config
    if (-not $Config.Pinned) { Push-To-Background }
    $miPin.Text = if ($Config.Pinned) { "Unpin (Allow Background)" } else { "Pin (Always On Top)" }
})

$miRefresh = New-Object System.Windows.Forms.ToolStripMenuItem
$miRefresh.Text = "Refresh Now"
$miRefresh.Add_Click({ Update-StorageInfo | Out-Null })

$miExit = New-Object System.Windows.Forms.ToolStripMenuItem
$miExit.Text = "Exit"
$miExit.Add_Click({
    try {
        $Config.Left = $window.Left
        $Config.Top  = $window.Top
        Save-Config
    } catch {}

    try { $wpfTimer.Stop() } catch {}
    try { $dispatcherPump.Stop() } catch {}
    try { $script:SaveTimer.Stop() } catch {}

    try { $window.Hide() } catch {}
    $script:AllowClose = $true

    try { $window.Dispatcher.Invoke([Action]{ $window.Close() }) } catch {}

    try {
        $notifyIcon.Visible = $false
        $notifyIcon.Dispose()
    } catch {}

    try { [System.Windows.Forms.Application]::Exit() } catch {}
})

$menu.Items.Add($miShowHide) | Out-Null
$menu.Items.Add($miPin)      | Out-Null
$menu.Items.Add($miRefresh)  | Out-Null
$menu.Items.Add((New-Object System.Windows.Forms.ToolStripSeparator)) | Out-Null
$menu.Items.Add($miExit)     | Out-Null
$notifyIcon.ContextMenuStrip = $menu

$notifyIcon.Add_DoubleClick({
    if ($window.IsVisible) {
        $window.Hide()
        $miShowHide.Text = "Show Widget"
    } else {
        $window.Show()
        $window.Activate() | Out-Null
        Maybe-PushToBackground
        $miShowHide.Text = "Hide Widget"
    }
})

# -----------------------------
# Show on launch
# -----------------------------
$window.Show()
$window.Activate() | Out-Null
Maybe-PushToBackground

# -----------------------------
# Keep WPF responsive while Forms runs
# IMPORTANT: Use BeginInvoke to avoid deadlocks during DragMove()
# -----------------------------
$dispatcherPump = New-Object System.Windows.Threading.DispatcherTimer
$dispatcherPump.Interval = [TimeSpan]::FromMilliseconds(200)
$dispatcherPump.Add_Tick({
    try {
        $null = $window.Dispatcher.BeginInvoke([Action]{}, [System.Windows.Threading.DispatcherPriority]::Background)
    } catch {}
})
$dispatcherPump.Start()

[System.Windows.Forms.Application]::Run()
