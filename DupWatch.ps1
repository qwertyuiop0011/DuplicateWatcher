<#
.SYNOPSIS
    DupWatch - catches duplicate downloads such as "photo (1).png" in the background.

.DESCRIPTION
    Run with no arguments to open the interactive menu (TUI):
        .\DupWatch.ps1

    Or control it directly from PowerShell:
        .\DupWatch.ps1 start                   turn ON (runs hidden in the background)
        .\DupWatch.ps1 stop                    turn OFF
        .\DupWatch.ps1 status                  show ON/OFF state and settings
        .\DupWatch.ps1 mode move|delete        move duplicates away, or delete them
        .\DupWatch.ps1 dupdir "D:\Dupes"       folder for moved duplicates
        .\DupWatch.ps1 watch "D:\MyDL"         folder to watch (default: Downloads)
        .\DupWatch.ps1 autostart on|off        start automatically at Windows login
        .\DupWatch.ps1 log [-Tail 20]          show recent activity

    Works in Windows PowerShell 5.1 and PowerShell 7.
#>
[CmdletBinding()]
param(
    [Parameter(Position = 0)]
    [ValidateSet('menu', 'start', 'stop', 'status', 'mode', 'dupdir', 'watch', 'autostart', 'log', 'run')]
    [string]$Command = 'menu',

    [Parameter(Position = 1)]
    [string]$Value,

    [int]$Tail = 20
)

# =====================================================================
#  Settings
# =====================================================================

$script:ScriptPath = $PSCommandPath
$script:AppDir     = Join-Path $env:APPDATA 'DupWatch'
$script:ConfigPath = Join-Path $script:AppDir 'config.json'
$script:LogPath    = Join-Path $script:AppDir 'dupwatch.log'
$script:MutexName  = 'Local\DupWatch_Running'
$script:EventName  = 'Local\DupWatch_Stop'
$script:RunKey     = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Run'
$script:RunValue   = 'DupWatch'
$script:PollMs     = 2000
$script:TempExt    = @('.crdownload', '.part', '.partial', '.tmp', '.download', '.opdownload')
$script:Downloads  = $null

[void][System.IO.Directory]::CreateDirectory($script:AppDir)

# =====================================================================
#  Config and log
# =====================================================================

function Get-DownloadsFolder {
    if ($script:Downloads) { return $script:Downloads }
    $path = $null
    try { $path = (New-Object -ComObject Shell.Application).NameSpace('shell:Downloads').Self.Path } catch {}
    if (-not $path) { $path = Join-Path $env:USERPROFILE 'Downloads' }
    $script:Downloads = $path
    $path
}

function Read-Config {
    $dl  = Get-DownloadsFolder
    $cfg = [pscustomobject]@{
        WatchFolder     = $dl
        DuplicateFolder = Join-Path $dl 'Duplicates'
        Mode            = 'Move'
    }
    if (Test-Path -LiteralPath $script:ConfigPath) {
        try {
            $saved = Get-Content -LiteralPath $script:ConfigPath -Raw -Encoding UTF8 | ConvertFrom-Json
            foreach ($k in 'WatchFolder', 'DuplicateFolder', 'Mode') { if ($saved.$k) { $cfg.$k = $saved.$k } }
        } catch {}
    }
    if ($cfg.Mode -ne 'Delete') { $cfg.Mode = 'Move' } else { $cfg.Mode = 'Delete' }
    $cfg
}

function Save-Config($cfg) {
    $cfg | ConvertTo-Json | Set-Content -LiteralPath $script:ConfigPath -Encoding UTF8
}

function Write-DupLog([string]$Message) {
    $line = '{0:yyyy-MM-dd HH:mm:ss}  {1}' -f (Get-Date), $Message
    try { Add-Content -LiteralPath $script:LogPath -Value $line -Encoding UTF8 } catch {}
}

function Read-DupLog([int]$Lines = 20) {
    if (-not (Test-Path -LiteralPath $script:LogPath)) { return @() }
    @(Get-Content -LiteralPath $script:LogPath -Tail $Lines -Encoding UTF8)
}

# =====================================================================
#  Duplicate detection
# =====================================================================

# "photo (1).png" -> "<same folder>\photo.png"; nothing if it isn't a duplicate-style name
function Get-OriginalPath([System.IO.FileInfo]$File) {
    if ($File.BaseName -match '^(.+) \((\d+)\)$') {
        Join-Path $File.DirectoryName ($Matches[1] + $File.Extension)
    }
}

function Test-FileLocked([string]$Path) {
    try {
        $fs = [System.IO.File]::Open($Path, 'Open', 'Read', 'None')
        $fs.Close()
        $false
    } catch { $true }
}

function Get-FileSet([string]$Folder) {
    $set = @{}
    Get-ChildItem -LiteralPath $Folder -File -Force -ErrorAction SilentlyContinue |
        ForEach-Object { $set[$_.FullName] = $true }
    $set
}

# Returns $true if the file was moved or deleted.
function Invoke-DuplicateAction([System.IO.FileInfo]$File, $Cfg) {
    $orig = Get-OriginalPath $File
    if (-not $orig -or -not (Test-Path -LiteralPath $orig -PathType Leaf)) { return $false }

    if ($Cfg.Mode -eq 'Delete') {
        $o    = Get-Item -LiteralPath $orig
        $same = ($o.Length -eq $File.Length) -and
                ((Get-FileHash -LiteralPath $orig).Hash -eq (Get-FileHash -LiteralPath $File.FullName).Hash)
        if ($same) {
            try {
                [System.IO.File]::Delete($File.FullName)
                Write-DupLog "Deleted duplicate: $($File.Name)"
                return $true
            } catch {
                Write-DupLog "Could not delete $($File.Name): $($_.Exception.Message)"
                return $false
            }
        }
        # Same name but different contents: never destroy it, move it instead.
        Write-DupLog "Contents differ from $($o.Name), moving instead of deleting: $($File.Name)"
    }

    try {
        [void][System.IO.Directory]::CreateDirectory($Cfg.DuplicateFolder)
        $dest = Join-Path $Cfg.DuplicateFolder $File.Name
        $i = 2
        while (Test-Path -LiteralPath $dest) {
            $dest = Join-Path $Cfg.DuplicateFolder ('{0} [{1}]{2}' -f $File.BaseName, $i, $File.Extension)
            $i++
        }
        [System.IO.File]::Move($File.FullName, $dest)   # also works across drives
        Write-DupLog "Moved duplicate: $($File.Name) -> $dest"
        $true
    } catch {
        Write-DupLog "Could not move $($File.Name): $($_.Exception.Message)"
        $false
    }
}

# One scan of the watched folder. $State holds what the watcher remembers between scans.
function Invoke-WatchPass($State) {
    $cfg = Read-Config   # settings changed from the menu or command line apply immediately

    if ($cfg.WatchFolder -ne $State.Watching) {
        $State.Watching = $cfg.WatchFolder
        $State.Seen     = Get-FileSet $State.Watching
        $State.Pending  = @{}
        Write-DupLog "Now watching $($State.Watching)"
        return
    }

    $files   = @(Get-ChildItem -LiteralPath $State.Watching -File -Force -ErrorAction SilentlyContinue)
    $current = @{}
    foreach ($f in $files) { $current[$f.FullName] = $true }

    # Forget files that disappeared, so a later re-download is noticed again.
    foreach ($k in @($State.Seen.Keys))    { if (-not $current.ContainsKey($k)) { $State.Seen.Remove($k) } }
    foreach ($k in @($State.Pending.Keys)) { if (-not $current.ContainsKey($k)) { $State.Pending.Remove($k) } }

    foreach ($f in $files) {
        if ($State.Seen.ContainsKey($f.FullName)) { continue }
        if ($script:TempExt -contains $f.Extension.ToLower()) { continue }
        if ($f.Length -eq 0) { continue }   # browser placeholder, still downloading

        # Wait until size/time hold still for one scan and nobody has the file open.
        $sig  = '{0}|{1}' -f $f.Length, $f.LastWriteTimeUtc.Ticks
        $busy = (Test-Path -LiteralPath "$($f.FullName).part") -or
                (Test-Path -LiteralPath "$($f.FullName).crdownload") -or
                (Test-FileLocked $f.FullName)
        if ($State.Pending[$f.FullName] -ne $sig -or $busy) {
            $State.Pending[$f.FullName] = $sig
            continue
        }

        $State.Pending.Remove($f.FullName)
        $State.Seen[$f.FullName] = $true
        try {
            if (Invoke-DuplicateAction $f $cfg) { $State.Seen.Remove($f.FullName) }
        } catch {
            Write-DupLog "Error on $($f.Name): $($_.Exception.Message)"
        }
    }
}

function New-WatchState {
    $cfg = Read-Config
    @{
        Watching = $cfg.WatchFolder
        Seen     = Get-FileSet $cfg.WatchFolder   # files already there are never touched
        Pending  = @{}
    }
}

# The hidden background process started by "start".
function Invoke-WatcherLoop {
    $created = $false
    $mutex = New-Object System.Threading.Mutex($true, $script:MutexName, [ref]$created)
    if (-not $created) { $mutex.Dispose(); return }   # already running

    $stop = New-Object System.Threading.EventWaitHandle($false, [System.Threading.EventResetMode]::ManualReset, $script:EventName)
    [void]$stop.Reset()
    try {
        $state = New-WatchState
        Write-DupLog "Started, watching $($state.Watching) (mode: $((Read-Config).Mode))"
        while (-not $stop.WaitOne($script:PollMs)) {
            try { Invoke-WatchPass $state } catch { Write-DupLog "Unexpected error: $($_.Exception.Message)" }
        }
        Write-DupLog 'Stopped'
    } finally {
        $stop.Dispose()
        try { $mutex.ReleaseMutex() } catch {}
        $mutex.Dispose()
    }
}

# =====================================================================
#  Control: on/off, settings, autostart (used by both the CLI and the TUI)
# =====================================================================

function Test-Running {
    $m = $null
    if ([System.Threading.Mutex]::TryOpenExisting($script:MutexName, [ref]$m)) { $m.Dispose(); return $true }
    $false
}

function Get-LaunchArgs([string]$Cmd) {
    $exe = (Get-Process -Id $PID).Path   # powershell.exe or pwsh.exe, whichever is running this
    [pscustomobject]@{
        Exe  = $exe
        Args = "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$($script:ScriptPath)`" $Cmd"
    }
}

function Wait-Running([bool]$Want, [int]$Seconds = 6) {
    $deadline = (Get-Date).AddSeconds($Seconds)
    while ((Get-Date) -lt $deadline) {
        if ((Test-Running) -eq $Want) { return $true }
        Start-Sleep -Milliseconds 200
    }
    (Test-Running) -eq $Want
}

function Start-Watcher {
    if (Test-Running) { return 'DupWatch is already ON.' }
    $l = Get-LaunchArgs 'run'
    Start-Process -FilePath $l.Exe -ArgumentList $l.Args -WindowStyle Hidden
    if (-not (Wait-Running $true 10)) { throw 'DupWatch failed to start. Check the log.' }
    "DupWatch is ON. Watching: $((Read-Config).WatchFolder)"
}

function Stop-Watcher {
    if (-not (Test-Running)) { return 'DupWatch is already OFF.' }
    $ev = $null
    if ([System.Threading.EventWaitHandle]::TryOpenExisting($script:EventName, [ref]$ev)) {
        [void]$ev.Set()
        $ev.Dispose()
    }
    if (-not (Wait-Running $false)) { throw 'DupWatch did not stop.' }
    'DupWatch is OFF.'
}

function Resolve-UserPath([string]$Path) {
    $p = "$Path".Trim().Trim('"').Trim("'")   # drag-and-dropped paths come with quotes
    if (-not $p) { throw 'No folder given.' }
    $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($p)
}

function Set-Setting([string]$Name, [string]$NewValue) {
    $cfg = Read-Config
    switch ($Name) {
        'mode' {
            if ($NewValue -notin 'move', 'delete') { throw "Mode must be 'move' or 'delete'." }
            if ($NewValue -eq 'delete') { $cfg.Mode = 'Delete' } else { $cfg.Mode = 'Move' }
        }
        'watch' {
            $p = Resolve-UserPath $NewValue
            if (-not (Test-Path -LiteralPath $p -PathType Container)) { throw "Folder does not exist: $p" }
            $cfg.WatchFolder = $p
        }
        'dupdir' { $cfg.DuplicateFolder = Resolve-UserPath $NewValue }
    }
    if ($cfg.DuplicateFolder.TrimEnd('\') -eq $cfg.WatchFolder.TrimEnd('\')) {
        throw 'The duplicate folder must be different from the watched folder.'
    }
    Save-Config $cfg
    $cfg
}

function Test-Autostart {
    $null -ne (Get-ItemProperty -Path $script:RunKey -Name $script:RunValue -ErrorAction SilentlyContinue)
}

function Set-Autostart([bool]$On) {
    if ($On) {
        $l = Get-LaunchArgs 'run'
        Set-ItemProperty -Path $script:RunKey -Name $script:RunValue -Value "`"$($l.Exe)`" $($l.Args)"
        'DupWatch will start automatically when you log in.'
    } else {
        Remove-ItemProperty -Path $script:RunKey -Name $script:RunValue -ErrorAction SilentlyContinue
        'DupWatch will no longer start at login.'
    }
}

function Get-Status {
    $cfg = Read-Config
    [pscustomobject]@{
        State           = $(if (Test-Running) { 'ON' } else { 'OFF' })
        Mode            = $cfg.Mode
        WatchFolder     = $cfg.WatchFolder
        DuplicateFolder = $cfg.DuplicateFolder
        Autostart       = $(if (Test-Autostart) { 'on' } else { 'off' })
        LogFile         = $script:LogPath
    }
}

# =====================================================================
#  TUI
# =====================================================================

# Box-drawing characters built from code points, so the file stays plain ASCII
# and displays correctly whatever encoding Windows PowerShell reads it with.
$script:G = @{
    H  = [string][char]0x2500; V  = [string][char]0x2502
    TL = [string][char]0x250C; TR = [string][char]0x2510
    BL = [string][char]0x2514; BR = [string][char]0x2518
    LT = [string][char]0x251C; RT = [string][char]0x2524
    Dot = [string][char]0x25CF; Up = [string][char]0x2191; Down = [string][char]0x2193
}

function Format-Fit([string]$Text, [int]$Width, [switch]$KeepEnd) {
    if ($null -eq $Text) { $Text = '' }
    if ($Width -le 0) { return '' }
    if ($Text.Length -le $Width) { return $Text.PadRight($Width) }
    if ($Width -le 3) { return $Text.Substring(0, $Width) }
    if ($KeepEnd) { return '...' + $Text.Substring($Text.Length - ($Width - 3)) }
    $Text.Substring(0, $Width - 3) + '...'
}

function Write-Seg([string]$Text, [ConsoleColor]$Fg = 'Gray', $Bg = $null) {
    if ($null -ne $Bg) { Write-Host $Text -NoNewline -ForegroundColor $Fg -BackgroundColor $Bg }
    else { Write-Host $Text -NoNewline -ForegroundColor $Fg }
}

function Write-Rule([string]$Left, [string]$Right, [int]$W, [string]$Title = '') {
    if ($Title) {
        $fill = $W - 4 - $Title.Length
        Write-Seg ($Left + $G.H + ' ') 'DarkGray'
        Write-Seg $Title 'Cyan'
        Write-Seg (' ' + ($G.H * [Math]::Max(0, $fill - 1)) + $Right) 'DarkGray'
    } else {
        Write-Seg ($Left + ($G.H * ($W - 2)) + $Right) 'DarkGray'
    }
    Write-Host ''
}

# A boxed row made of colored segments: @( @(text, color [, background]), ... )
function Write-Row([int]$W, [object[]]$Segments) {
    $inner = $W - 4
    Write-Seg ($G.V + ' ') 'DarkGray'
    $used = 0
    foreach ($s in $Segments) {
        $room = $inner - $used
        if ($room -le 0) { break }
        $t = $s[0]
        if ($t.Length -gt $room) { $t = $t.Substring(0, $room) }
        if ($s.Count -ge 3) { Write-Seg $t $s[1] $s[2] } else { Write-Seg $t $s[1] }
        $used += $t.Length
    }
    if ($used -lt $inner) { Write-Seg (' ' * ($inner - $used)) }
    Write-Seg (' ' + $G.V) 'DarkGray'
    Write-Host ''
}

function Get-MenuItems($St) {
    @(
        $(if ($St.State -eq 'ON') { 'Turn OFF' } else { 'Turn ON' }),
        $(if ($St.Mode -eq 'Delete') { 'Switch to Move mode' } else { 'Switch to Delete mode' }),
        'Change watched folder',
        'Change duplicate folder',
        $(if ($St.Autostart -eq 'on') { 'Turn autostart OFF' } else { 'Turn autostart ON' }),
        'View full log',
        'Open duplicate folder',
        'Quit'
    )
}

function Show-Screen([int]$Sel, [string]$Msg, $St, [int]$W) {
    [Console]::SetCursorPosition(0, 0)
    $inner = $W - 4
    $label = 13

    Write-Rule $G.TL $G.TR $W 'DupWatch'
    if ($St.State -eq 'ON') { $stColor = 'Green' } else { $stColor = 'Red' }
    Write-Row $W @(@((Format-Fit 'Status' $label), 'DarkGray'), @("$($G.Dot) $($St.State)", $stColor))
    if ($St.Mode -eq 'Delete') { $modeText = 'Delete exact duplicates' } else { $modeText = 'Move duplicates to the duplicate folder' }
    Write-Row $W @(@((Format-Fit 'Mode' $label), 'DarkGray'), @($modeText, 'White'))
    Write-Row $W @(@((Format-Fit 'Watching' $label), 'DarkGray'), @((Format-Fit $St.WatchFolder ($inner - $label) -KeepEnd), 'White'))
    Write-Row $W @(@((Format-Fit 'Dup folder' $label), 'DarkGray'), @((Format-Fit $St.DuplicateFolder ($inner - $label) -KeepEnd), 'White'))
    Write-Row $W @(@((Format-Fit 'Autostart' $label), 'DarkGray'), @($St.Autostart, 'White'))
    Write-Rule $G.LT $G.RT $W

    $items = Get-MenuItems $St
    for ($i = 0; $i -lt $items.Count; $i++) {
        $text = Format-Fit (" {0}  {1}" -f ($i + 1), $items[$i]) $inner
        if ($i -eq $Sel) { Write-Row $W @(, @($text, 'Black', 'Cyan')) }
        else { Write-Row $W @(, @($text, 'Gray')) }
    }

    Write-Rule $G.LT $G.RT $W 'Recent activity'
    $log = @(Read-DupLog 5)
    for ($i = 0; $i -lt 5; $i++) {
        if ($i -lt $log.Count) {
            $line = $log[$i]
            if ($line.Length -gt 11) { $line = $line.Substring(11) }   # drop the date, keep the time
            if ($line -match 'Deleted') { $c = 'Yellow' } elseif ($line -match 'Moved') { $c = 'Cyan' }
            elseif ($line -match 'Could not|error') { $c = 'Red' } else { $c = 'DarkGray' }
            Write-Row $W @(, @((Format-Fit $line $inner), $c))
        } elseif ($i -eq 0) {
            Write-Row $W @(, @('No activity yet.', 'DarkGray'))
        } else {
            Write-Row $W @(, @('', 'Gray'))
        }
    }
    Write-Rule $G.BL $G.BR $W

    Write-Seg (Format-Fit " $($G.Up)$($G.Down) move   Enter select   1-8 shortcut   Q quit" ($W)) 'DarkGray'
    Write-Host ''
    Write-Seg (Format-Fit " $Msg" $W) 'Yellow'
    Write-Host ''
}

# Simple line editor on top of ReadKey (same input method as the menu, so the two never fight).
# Enter = accept, Esc = cancel (returns $null), Backspace = delete.
function Read-TuiLine([string]$Prompt) {
    Write-Host $Prompt -NoNewline -ForegroundColor Cyan
    $buf = New-Object System.Text.StringBuilder
    while ($true) {
        $k = [Console]::ReadKey($true)
        switch ($k.Key) {
            'Enter'  { Write-Host ''; return $buf.ToString() }
            'Escape' { Write-Host ''; return $null }
            'Backspace' {
                if ($buf.Length -gt 0) { [void]$buf.Remove($buf.Length - 1, 1); Write-Host "`b `b" -NoNewline }
            }
            default {
                if (-not [char]::IsControl($k.KeyChar)) { [void]$buf.Append($k.KeyChar); Write-Host $k.KeyChar -NoNewline }
            }
        }
    }
}

function Read-FolderInput([string]$What, [string]$Current) {
    Clear-Host
    [Console]::CursorVisible = $true
    Write-Host ''
    Write-Host "  Change $What" -ForegroundColor Cyan
    Write-Host "  Current: $Current" -ForegroundColor DarkGray
    Write-Host ''
    Write-Host '  Type or drag a folder into this window, then press Enter.'
    Write-Host '  Esc or an empty line cancels.' -ForegroundColor DarkGray
    Write-Host ''
    try { $answer = Read-TuiLine '  > ' }
    finally { [Console]::CursorVisible = $false; Clear-Host }
    $answer
}

function Show-FullLog {
    Clear-Host
    $rows = [Math]::Max(5, [Console]::WindowHeight - 5)
    Write-Host ''
    Write-Host "  Activity log  ($($script:LogPath))" -ForegroundColor Cyan
    Write-Host ''
    $lines = @(Read-DupLog $rows)
    if ($lines.Count -eq 0) { Write-Host '  No activity yet.' -ForegroundColor DarkGray }
    foreach ($l in $lines) { Write-Host "  $l" }
    Write-Host ''
    Write-Host '  Press any key to go back...' -ForegroundColor DarkGray
    [void][Console]::ReadKey($true)
    Clear-Host
}

function Invoke-MenuAction([int]$Index, $St) {
    switch ($Index) {
        0 { if ($St.State -eq 'ON') { Stop-Watcher } else { Start-Watcher } }
        1 {
            if ($St.Mode -eq 'Delete') { [void](Set-Setting 'mode' 'move'); 'Mode: move duplicates.' }
            else { [void](Set-Setting 'mode' 'delete'); 'Mode: delete exact duplicates (different contents are still moved).' }
        }
        2 {
            $a = Read-FolderInput 'watched folder' $St.WatchFolder
            if ("$a".Trim()) { $c = Set-Setting 'watch' $a; "Now watching $($c.WatchFolder)" } else { 'Cancelled.' }
        }
        3 {
            $a = Read-FolderInput 'duplicate folder' $St.DuplicateFolder
            if ("$a".Trim()) { $c = Set-Setting 'dupdir' $a; "Duplicates go to $($c.DuplicateFolder)" } else { 'Cancelled.' }
        }
        4 { Set-Autostart ($St.Autostart -ne 'on') }
        5 { Show-FullLog; '' }
        6 {
            [void][System.IO.Directory]::CreateDirectory($St.DuplicateFolder)
            Invoke-Item -LiteralPath $St.DuplicateFolder
            'Opened the duplicate folder.'
        }
    }
}

function Show-Menu {
    $sel = 0
    $msg = 'Welcome! Use the arrow keys and Enter.'
    $count = 8
    $lastDraw = [datetime]::MinValue
    $lastWidth = -1
    $dirty = $true

    $oldTitle = $Host.UI.RawUI.WindowTitle
    $Host.UI.RawUI.WindowTitle = 'DupWatch'
    [Console]::CursorVisible = $false
    Clear-Host
    try {
        while ($true) {
            $winW = [Console]::WindowWidth
            $W = [Math]::Max(46, [Math]::Min(80, $winW - 1))
            if ($winW -ne $lastWidth) { Clear-Host; $lastWidth = $winW; $dirty = $true }

            # Redraw after a key press, and once a second so status and log stay live.
            if ($dirty -or ((Get-Date) - $lastDraw).TotalMilliseconds -ge 1000) {
                $st = Get-Status
                Show-Screen $sel $msg $st $W
                $lastDraw = Get-Date
                $dirty = $false
            }

            if (-not [Console]::KeyAvailable) { Start-Sleep -Milliseconds 80; continue }
            $key = [Console]::ReadKey($true)
            $dirty = $true
            $run = -1
            switch ($key.Key) {
                'UpArrow'   { $sel = ($sel + $count - 1) % $count }
                'DownArrow' { $sel = ($sel + 1) % $count }
                'Home'      { $sel = 0 }
                'End'       { $sel = $count - 1 }
                'Enter'     { $run = $sel }
                'Escape'    { $run = $count - 1 }
                'Q'         { $run = $count - 1 }
                default {
                    $n = 0
                    if ([int]::TryParse([string]$key.KeyChar, [ref]$n) -and $n -ge 1 -and $n -le $count) {
                        $sel = $n - 1
                        $run = $sel
                    }
                }
            }
            if ($run -lt 0) { continue }
            if ($run -eq $count - 1) { break }   # Quit

            if ($run -eq 0) {   # starting/stopping takes a moment: say so first
                if ($st.State -eq 'ON') { $msg = 'Stopping...' } else { $msg = 'Starting...' }
                Show-Screen $sel $msg $st $W
            }
            try { $msg = [string](Invoke-MenuAction $run $st | Select-Object -Last 1) }
            catch { $msg = "Error: $($_.Exception.Message)" }
            $lastWidth = -1   # force a clean redraw after prompts/sub-screens
        }
    } finally {
        [Console]::CursorVisible = $true
        $Host.UI.RawUI.WindowTitle = $oldTitle
        Clear-Host
        if (Test-Running) { Write-Host 'DupWatch is still ON in the background.' -ForegroundColor Green }
        else { Write-Host 'DupWatch is OFF.' -ForegroundColor DarkGray }
    }
}

# =====================================================================
#  Entry point
# =====================================================================

try {
    switch ($Command) {
        'run'    { Invoke-WatcherLoop }
        'menu'   { Show-Menu }
        'start'  { Start-Watcher }
        'stop'   { Stop-Watcher }
        'status' {
            $s = Get-Status
            "State:            $($s.State)"
            "Mode:             $($s.Mode)"
            "Watching:         $($s.WatchFolder)"
            "Duplicate folder: $($s.DuplicateFolder)"
            "Autostart:        $($s.Autostart)"
            "Log file:         $($s.LogFile)"
        }
        'log'    { $l = Read-DupLog $Tail; if ($l) { $l } else { 'No activity yet.' } }
        'autostart' {
            if ($Value -notin 'on', 'off') { throw "Use: autostart on|off" }
            Set-Autostart ($Value -eq 'on')
        }
        { $_ -in 'mode', 'watch', 'dupdir' } {
            if (-not $Value) { throw "Use: $Command <value>" }
            $c = Set-Setting $Command $Value
            switch ($Command) {
                'mode'   { "Mode set to $($c.Mode)." }
                'watch'  { "Watch folder set to $($c.WatchFolder)" }
                'dupdir' { "Duplicate folder set to $($c.DuplicateFolder)" }
            }
        }
    }
} catch {
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}
