<#
.SYNOPSIS
    DupWatch - "사진 (1).png" 처럼 중복으로 다운로드된 파일을 백그라운드에서 자동으로 처리합니다.

.DESCRIPTION
    인수 없이 실행하면 대화형 메뉴(TUI)가 열립니다:
        .\DupWatch.ps1

    PowerShell 명령으로 직접 조작할 수도 있습니다:
        .\DupWatch.ps1 start                   켜기 (백그라운드에서 숨겨진 채로 실행)
        .\DupWatch.ps1 stop                    끄기
        .\DupWatch.ps1 status                  켜짐/꺼짐 상태와 설정 보기
        .\DupWatch.ps1 mode move|delete        중복 파일을 중복 폴더로 이동 / 삭제
        .\DupWatch.ps1 keep oldest|newest      가장 오래된 파일 유지 / 가장 최근 파일 유지
        .\DupWatch.ps1 dupdir "D:\중복"         중복 파일을 옮길 폴더
        .\DupWatch.ps1 watch "D:\다운로드"      감시할 폴더 (기본값: 다운로드 폴더)
        .\DupWatch.ps1 autostart on|off        Windows 로그인 시 자동 시작
        .\DupWatch.ps1 log [-Tail 20]          최근 활동 기록 보기

    보관 기준 (keep):
        oldest  기존 파일("사진.png")을 그대로 두고, 새로 받은 "사진 (1).png" 를 처리합니다. (기본값)
        newest  기존 "사진.png" 를 처리하고, 새로 받은 파일의 이름을 "사진.png" 로 바꿔 남깁니다.

    Windows PowerShell 5.1 과 PowerShell 7 에서 모두 동작합니다.
    이 파일은 BOM 이 있는 UTF-8 로 저장되어 있어야 한글이 깨지지 않습니다.
#>
[CmdletBinding()]
param(
    [Parameter(Position = 0)]
    [ValidateSet('menu', 'start', 'stop', 'status', 'mode', 'keep', 'dupdir', 'watch', 'autostart', 'log', 'run')]
    [string]$Command = 'menu',

    [Parameter(Position = 1)]
    [string]$Value,

    [int]$Tail = 20
)

# =====================================================================
#  기본 설정
# =====================================================================

$script:ScriptPath = $PSCommandPath
$script:AppDir     = Join-Path $env:APPDATA 'DupWatch'
$script:ConfigPath = Join-Path $script:AppDir 'config.json'
$script:LogPath    = Join-Path $script:AppDir 'dupwatch.log'
$script:MutexName  = 'Local\DupWatch_Running'   # 이 이름의 뮤텍스가 있으면 "켜짐"
$script:EventName  = 'Local\DupWatch_Stop'      # 이 이벤트를 신호하면 "꺼짐"
$script:RunKey     = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Run'
$script:RunValue   = 'DupWatch'
$script:PollMs     = 2000                       # 폴더 검사 간격 (밀리초)
$script:TempExt    = @('.crdownload', '.part', '.partial', '.tmp', '.download', '.opdownload')
$script:Downloads  = $null

[void][System.IO.Directory]::CreateDirectory($script:AppDir)

# =====================================================================
#  설정 파일과 기록 파일
# =====================================================================

# 실제 "다운로드" 폴더 위치 (다른 드라이브로 옮겨 둔 경우도 찾아냄)
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
        Mode            = 'Move'      # Move = 이동, Delete = 삭제
        Keep            = 'Oldest'    # Oldest = 오래된 파일 유지, Newest = 최근 파일 유지
    }
    if (Test-Path -LiteralPath $script:ConfigPath) {
        try {
            $saved = Get-Content -LiteralPath $script:ConfigPath -Raw -Encoding UTF8 | ConvertFrom-Json
            foreach ($k in 'WatchFolder', 'DuplicateFolder', 'Mode', 'Keep') { if ($saved.$k) { $cfg.$k = $saved.$k } }
        } catch {}
    }
    # 잘못된 값이 들어 있으면 기본값으로 되돌림
    if ($cfg.Mode -eq 'Delete') { $cfg.Mode = 'Delete' } else { $cfg.Mode = 'Move' }
    if ($cfg.Keep -eq 'Newest') { $cfg.Keep = 'Newest' } else { $cfg.Keep = 'Oldest' }
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
#  중복 파일 감지와 처리
# =====================================================================

# "사진 (1).png" -> "<같은 폴더>\사진.png". 중복 형식의 이름이 아니면 아무것도 반환하지 않음
function Get-OriginalPath([System.IO.FileInfo]$File) {
    if ($File.BaseName -match '^(.+) \((\d+)\)$') {
        Join-Path $File.DirectoryName ($Matches[1] + $File.Extension)
    }
}

# 다른 프로그램(브라우저 등)이 아직 파일을 열고 있으면 $true
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

function Test-SameContent([System.IO.FileInfo]$A, [System.IO.FileInfo]$B) {
    ($A.Length -eq $B.Length) -and
        ((Get-FileHash -LiteralPath $A.FullName).Hash -eq (Get-FileHash -LiteralPath $B.FullName).Hash)
}

# $Target 파일을 없앰: 삭제 모드이고 $Kept 와 내용이 똑같으면 삭제, 그 밖에는 중복 폴더로 이동.
# 성공하면 $true.
function Remove-DuplicateFile([System.IO.FileInfo]$Target, [System.IO.FileInfo]$Kept, $Cfg) {
    if ($Cfg.Mode -eq 'Delete') {
        if (Test-SameContent $Target $Kept) {
            try {
                [System.IO.File]::Delete($Target.FullName)
                Write-DupLog "중복 파일 삭제: $($Target.Name)"
                return $true
            } catch {
                Write-DupLog "삭제 실패 $($Target.Name): $($_.Exception.Message)"
                return $false
            }
        }
        # 이름만 같고 내용이 다르면 절대 지우지 않고 이동시킴
        Write-DupLog "내용이 $($Kept.Name) 와(과) 달라 삭제 대신 이동: $($Target.Name)"
    }

    try {
        [void][System.IO.Directory]::CreateDirectory($Cfg.DuplicateFolder)
        $dest = Join-Path $Cfg.DuplicateFolder $Target.Name
        $i = 2
        while (Test-Path -LiteralPath $dest) {
            $dest = Join-Path $Cfg.DuplicateFolder ('{0} [{1}]{2}' -f $Target.BaseName, $i, $Target.Extension)
            $i++
        }
        [System.IO.File]::Move($Target.FullName, $dest)   # 다른 드라이브로도 이동 가능
        Write-DupLog "중복 파일 이동: $($Target.Name) -> $dest"
        $true
    } catch {
        Write-DupLog "이동 실패 $($Target.Name): $($_.Exception.Message)"
        $false
    }
}

# 새로 받은 $File 이 기존 파일의 중복이면 설정에 따라 처리. 감시 폴더에서 $File 이 사라졌으면 $true.
function Invoke-DuplicateAction([System.IO.FileInfo]$File, $Cfg) {
    $orig = Get-OriginalPath $File
    if (-not $orig -or -not (Test-Path -LiteralPath $orig -PathType Leaf)) { return $false }   # 중복 아님
    $origFile = Get-Item -LiteralPath $orig

    if ($Cfg.Keep -eq 'Newest') {
        # 최근 파일 유지: 기존 파일을 처리하고, 새 파일에 원래 이름을 붙임
        if (-not (Remove-DuplicateFile $origFile $File $Cfg)) { return $false }
        try {
            [System.IO.File]::Move($File.FullName, $orig)
            Write-DupLog "최근 파일 유지: $($File.Name) -> $($origFile.Name) 로 이름 변경"
        } catch {
            Write-DupLog "이름 변경 실패 $($File.Name): $($_.Exception.Message)"
        }
        return $true
    }

    # 오래된 파일 유지 (기본값): 기존 파일은 그대로 두고 새 파일을 처리
    Remove-DuplicateFile $File $origFile $Cfg
}

# 감시 폴더를 한 번 검사. $State 에는 검사 사이에 기억해야 하는 내용이 들어 있음
function Invoke-WatchPass($State) {
    $cfg = Read-Config   # 메뉴나 명령으로 바꾼 설정이 바로 적용되도록 매번 다시 읽음

    if ($cfg.WatchFolder -ne $State.Watching) {
        $State.Watching = $cfg.WatchFolder
        $State.Seen     = Get-FileSet $State.Watching
        $State.Pending  = @{}
        Write-DupLog "감시 폴더 변경: $($State.Watching)"
        return
    }

    $files   = @(Get-ChildItem -LiteralPath $State.Watching -File -Force -ErrorAction SilentlyContinue)
    $current = @{}
    foreach ($f in $files) { $current[$f.FullName] = $true }

    # 사라진 파일은 잊어버림 (같은 이름으로 다시 받으면 다시 감지되도록)
    foreach ($k in @($State.Seen.Keys))    { if (-not $current.ContainsKey($k)) { $State.Seen.Remove($k) } }
    foreach ($k in @($State.Pending.Keys)) { if (-not $current.ContainsKey($k)) { $State.Pending.Remove($k) } }

    foreach ($f in $files) {
        if ($State.Seen.ContainsKey($f.FullName)) { continue }
        if ($script:TempExt -contains $f.Extension.ToLower()) { continue }   # 다운로드 중인 임시 파일
        if ($f.Length -eq 0) { continue }   # 브라우저가 만든 빈 자리표시 파일, 아직 받는 중

        # 크기와 수정 시각이 한 번의 검사 동안 그대로이고, 아무도 파일을 열고 있지 않을 때까지 기다림
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
            Write-DupLog "오류 $($f.Name): $($_.Exception.Message)"
        }
    }
}

function New-WatchState {
    $cfg = Read-Config
    @{
        Watching = $cfg.WatchFolder
        Seen     = Get-FileSet $cfg.WatchFolder   # 시작할 때 이미 있던 파일은 건드리지 않음
        Pending  = @{}
    }
}

# "start" 가 띄우는 숨겨진 백그라운드 프로세스의 본체
function Invoke-WatcherLoop {
    $created = $false
    $mutex = New-Object System.Threading.Mutex($true, $script:MutexName, [ref]$created)
    if (-not $created) { $mutex.Dispose(); return }   # 이미 실행 중

    $stop = New-Object System.Threading.EventWaitHandle($false, [System.Threading.EventResetMode]::ManualReset, $script:EventName)
    [void]$stop.Reset()
    try {
        $state = New-WatchState
        $c = Read-Config
        Write-DupLog "시작됨, 감시 폴더: $($state.Watching) (처리 방식: $(Get-ModeText $c.Mode), 보관 기준: $(Get-KeepText $c.Keep))"
        while (-not $stop.WaitOne($script:PollMs)) {
            try { Invoke-WatchPass $state } catch { Write-DupLog "예상치 못한 오류: $($_.Exception.Message)" }
        }
        Write-DupLog '중지됨'
    } finally {
        $stop.Dispose()
        try { $mutex.ReleaseMutex() } catch {}
        $mutex.Dispose()
    }
}

# =====================================================================
#  제어: 켜기/끄기, 설정, 자동 시작 (명령줄과 TUI 가 함께 사용)
# =====================================================================

function Get-ModeText([string]$Mode) { if ($Mode -eq 'Delete') { '삭제' } else { '이동' } }
function Get-KeepText([string]$Keep) { if ($Keep -eq 'Newest') { '최근 파일 유지' } else { '오래된 파일 유지' } }

function Test-Running {
    $m = $null
    if ([System.Threading.Mutex]::TryOpenExisting($script:MutexName, [ref]$m)) { $m.Dispose(); return $true }
    $false
}

function Get-LaunchArgs([string]$Cmd) {
    $exe = (Get-Process -Id $PID).Path   # 지금 실행 중인 powershell.exe 또는 pwsh.exe
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
    if (Test-Running) { return 'DupWatch 가 이미 켜져 있습니다.' }
    $l = Get-LaunchArgs 'run'
    Start-Process -FilePath $l.Exe -ArgumentList $l.Args -WindowStyle Hidden
    if (-not (Wait-Running $true 10)) { throw 'DupWatch 를 시작하지 못했습니다. 활동 기록을 확인하세요.' }
    "DupWatch 가 켜졌습니다. 감시 폴더: $((Read-Config).WatchFolder)"
}

function Stop-Watcher {
    if (-not (Test-Running)) { return 'DupWatch 가 이미 꺼져 있습니다.' }
    $ev = $null
    if ([System.Threading.EventWaitHandle]::TryOpenExisting($script:EventName, [ref]$ev)) {
        [void]$ev.Set()
        $ev.Dispose()
    }
    if (-not (Wait-Running $false)) { throw 'DupWatch 가 중지되지 않았습니다.' }
    'DupWatch 가 꺼졌습니다.'
}

function Resolve-UserPath([string]$Path) {
    $p = "$Path".Trim().Trim('"').Trim("'")   # 끌어다 놓은 경로에는 따옴표가 붙어 있음
    if (-not $p) { throw '폴더가 입력되지 않았습니다.' }
    $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($p)
}

function Set-Setting([string]$Name, [string]$NewValue) {
    $cfg = Read-Config
    switch ($Name) {
        'mode' {
            if ($NewValue -notin 'move', 'delete') { throw "처리 방식은 'move'(이동) 또는 'delete'(삭제)만 가능합니다." }
            if ($NewValue -eq 'delete') { $cfg.Mode = 'Delete' } else { $cfg.Mode = 'Move' }
        }
        'keep' {
            if ($NewValue -notin 'oldest', 'newest') { throw "보관 기준은 'oldest'(오래된 파일) 또는 'newest'(최근 파일)만 가능합니다." }
            if ($NewValue -eq 'newest') { $cfg.Keep = 'Newest' } else { $cfg.Keep = 'Oldest' }
        }
        'watch' {
            $p = Resolve-UserPath $NewValue
            if (-not (Test-Path -LiteralPath $p -PathType Container)) { throw "폴더가 존재하지 않습니다: $p" }
            $cfg.WatchFolder = $p
        }
        'dupdir' { $cfg.DuplicateFolder = Resolve-UserPath $NewValue }
    }
    if ($cfg.DuplicateFolder.TrimEnd('\') -eq $cfg.WatchFolder.TrimEnd('\')) {
        throw '중복 폴더는 감시 폴더와 달라야 합니다.'
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
        'Windows 에 로그인하면 DupWatch 가 자동으로 시작됩니다.'
    } else {
        Remove-ItemProperty -Path $script:RunKey -Name $script:RunValue -ErrorAction SilentlyContinue
        '이제 로그인할 때 DupWatch 가 자동으로 시작되지 않습니다.'
    }
}

function Get-Status {
    $cfg = Read-Config
    [pscustomobject]@{
        State           = $(if (Test-Running) { 'ON' } else { 'OFF' })
        Mode            = $cfg.Mode
        Keep            = $cfg.Keep
        WatchFolder     = $cfg.WatchFolder
        DuplicateFolder = $cfg.DuplicateFolder
        Autostart       = $(if (Test-Autostart) { 'on' } else { 'off' })
        LogFile         = $script:LogPath
    }
}

# =====================================================================
#  TUI (화면 메뉴)
# =====================================================================

# 테두리 문자. 한국어 Windows 의 기존 콘솔에서는 선 문자가 두 칸을 차지해 상자가 깨질 수 있으므로,
# 메뉴를 열 때 실제 폭을 재어 보고 두 칸이면 ASCII 문자(+ - |)로 바꿈 (Initialize-Glyphs).
$script:G = @{
    H  = [string][char]0x2500; V  = [string][char]0x2502
    TL = [string][char]0x250C; TR = [string][char]0x2510
    BL = [string][char]0x2514; BR = [string][char]0x2518
    LT = [string][char]0x251C; RT = [string][char]0x2524
    Dot = [string][char]0x25CF
}

# 콘솔에서 두 칸을 차지하는 문자 (한글, 한자, 일본어, 전각 문자)
$script:WideRx = New-Object System.Text.RegularExpressions.Regex '[\u1100-\u115F\u2E80-\uA4CF\uAC00-\uD7A3\uF900-\uFAFF\uFE30-\uFE4F\uFF00-\uFF60\uFFE0-\uFFE6]'

function Initialize-Glyphs {
    if ([Environment]::OSVersion.Platform -ne 'Win32NT') { return }
    $wide = $false
    try {
        [Console]::SetCursorPosition(0, 0)
        Write-Host $G.H -NoNewline
        $wide = [Console]::CursorLeft -ne 1
        [Console]::SetCursorPosition(0, 0)
        Write-Host '    ' -NoNewline
    } catch {}
    if ($wide) {
        $script:G = @{ H = '-'; V = '|'; TL = '+'; TR = '+'; BL = '+'; BR = '+'; LT = '+'; RT = '+'; Dot = '*' }
    }
}

# 문자열이 화면에서 차지하는 칸 수
function Get-TextWidth([string]$Text) {
    if (-not $Text) { return 0 }
    $Text.Length + $script:WideRx.Matches($Text).Count
}

function Get-CharWidth([char]$Ch) {
    if ($script:WideRx.IsMatch([string]$Ch)) { 2 } else { 1 }
}

# 정확히 $Width 칸이 되도록 자르거나 공백으로 채움. -KeepEnd 면 앞쪽을 "..." 으로 줄임 (경로용)
function Format-Fit([string]$Text, [int]$Width, [switch]$KeepEnd) {
    if ($null -eq $Text) { $Text = '' }
    if ($Width -le 0) { return '' }
    $tw = Get-TextWidth $Text
    if ($tw -le $Width) { return $Text + (' ' * ($Width - $tw)) }
    $room = $Width - 3
    if ($room -le 0) { return '.' * $Width }
    $sb = New-Object System.Text.StringBuilder
    $used = 0
    if ($KeepEnd) {
        for ($i = $Text.Length - 1; $i -ge 0; $i--) {
            $cw = Get-CharWidth $Text[$i]
            if ($used + $cw -gt $room) { break }
            [void]$sb.Insert(0, $Text[$i]); $used += $cw
        }
        return '...' + $sb.ToString() + (' ' * ($room - $used))
    }
    foreach ($ch in $Text.ToCharArray()) {
        $cw = Get-CharWidth $ch
        if ($used + $cw -gt $room) { break }
        [void]$sb.Append($ch); $used += $cw
    }
    $sb.ToString() + '...' + (' ' * ($room - $used))
}

function Write-Seg([string]$Text, [ConsoleColor]$Fg = 'Gray', $Bg = $null) {
    if ($null -ne $Bg) { Write-Host $Text -NoNewline -ForegroundColor $Fg -BackgroundColor $Bg }
    else { Write-Host $Text -NoNewline -ForegroundColor $Fg }
}

# 가로선. $Title 이 있으면 선 안에 제목을 넣음
function Write-Rule([string]$Left, [string]$Right, [int]$W, [string]$Title = '') {
    if ($Title) {
        $fill = $W - 5 - (Get-TextWidth $Title)
        Write-Seg ($Left + $G.H + ' ') 'DarkGray'
        Write-Seg $Title 'Cyan'
        Write-Seg (' ' + ($G.H * [Math]::Max(0, $fill)) + $Right) 'DarkGray'
    } else {
        Write-Seg ($Left + ($G.H * ($W - 2)) + $Right) 'DarkGray'
    }
    Write-Host ''
}

# 색이 다른 조각들로 이루어진 상자 안의 한 줄: @( @(글자, 글자색 [, 배경색]), ... )
function Write-Row([int]$W, [object[]]$Segments) {
    $inner = $W - 4
    Write-Seg ($G.V + ' ') 'DarkGray'
    $used = 0
    foreach ($s in $Segments) {
        $room = $inner - $used
        if ($room -le 0) { break }
        $t = [string]$s[0]
        if ((Get-TextWidth $t) -gt $room) { $t = Format-Fit $t $room }
        if ($s.Count -ge 3) { Write-Seg $t $s[1] $s[2] } else { Write-Seg $t $s[1] }
        $used += Get-TextWidth $t
    }
    if ($used -lt $inner) { Write-Seg (' ' * ($inner - $used)) }
    Write-Seg (' ' + $G.V) 'DarkGray'
    Write-Host ''
}

function Get-MenuItems($St) {
    @(
        $(if ($St.State -eq 'ON') { '끄기' } else { '켜기' }),
        $(if ($St.Mode -eq 'Delete') { '이동 모드로 전환' } else { '삭제 모드로 전환' }),
        $(if ($St.Keep -eq 'Newest') { '오래된 파일 유지로 전환' } else { '최근 파일 유지로 전환' }),
        '감시 폴더 변경',
        '중복 폴더 변경',
        $(if ($St.Autostart -eq 'on') { '자동 시작 끄기' } else { '자동 시작 켜기' }),
        '전체 기록 보기',
        '중복 폴더 열기',
        '종료'
    )
}

function Show-Screen([int]$Sel, [string]$Msg, $St, [int]$W) {
    [Console]::SetCursorPosition(0, 0)
    $inner = $W - 4
    $label = 12

    Write-Rule $G.TL $G.TR $W 'DupWatch'
    if ($St.State -eq 'ON') { $stText = "$($G.Dot) 켜짐"; $stColor = 'Green' } else { $stText = "$($G.Dot) 꺼짐"; $stColor = 'Red' }
    Write-Row $W @(@((Format-Fit '상태' $label), 'DarkGray'), @($stText, $stColor))
    if ($St.Mode -eq 'Delete') { $modeText = '삭제 (내용이 똑같은 중복 파일만 삭제)' } else { $modeText = '이동 (중복 파일을 중복 폴더로 이동)' }
    Write-Row $W @(@((Format-Fit '처리 방식' $label), 'DarkGray'), @($modeText, 'White'))
    if ($St.Keep -eq 'Newest') { $keepText = '최근 파일 유지 (새로 받은 파일을 남김)' } else { $keepText = '오래된 파일 유지 (기존 파일을 남김)' }
    Write-Row $W @(@((Format-Fit '보관 기준' $label), 'DarkGray'), @($keepText, 'White'))
    Write-Row $W @(@((Format-Fit '감시 폴더' $label), 'DarkGray'), @((Format-Fit $St.WatchFolder ($inner - $label) -KeepEnd), 'White'))
    Write-Row $W @(@((Format-Fit '중복 폴더' $label), 'DarkGray'), @((Format-Fit $St.DuplicateFolder ($inner - $label) -KeepEnd), 'White'))
    if ($St.Autostart -eq 'on') { $autoText = '켜짐' } else { $autoText = '꺼짐' }
    Write-Row $W @(@((Format-Fit '자동 시작' $label), 'DarkGray'), @($autoText, 'White'))
    Write-Rule $G.LT $G.RT $W

    $items = Get-MenuItems $St
    for ($i = 0; $i -lt $items.Count; $i++) {
        $text = Format-Fit (" {0}  {1}" -f ($i + 1), $items[$i]) $inner
        if ($i -eq $Sel) { Write-Row $W @(, @($text, 'Black', 'Cyan')) }
        else { Write-Row $W @(, @($text, 'Gray')) }
    }

    Write-Rule $G.LT $G.RT $W '최근 활동'
    $log = @(Read-DupLog 5)
    for ($i = 0; $i -lt 5; $i++) {
        if ($i -lt $log.Count) {
            $line = $log[$i]
            if ($line.Length -gt 11) { $line = $line.Substring(11) }   # 날짜는 빼고 시각만 표시
            if ($line -match '실패|오류') { $c = 'Red' } elseif ($line -match '삭제') { $c = 'Yellow' }
            elseif ($line -match '이동|유지') { $c = 'Cyan' } else { $c = 'DarkGray' }
            Write-Row $W @(, @((Format-Fit $line $inner), $c))
        } elseif ($i -eq 0) {
            Write-Row $W @(, @('아직 기록이 없습니다.', 'DarkGray'))
        } else {
            Write-Row $W @(, @('', 'Gray'))
        }
    }
    Write-Rule $G.BL $G.BR $W

    Write-Seg (Format-Fit ' 방향키 이동   Enter 선택   1-9 바로가기   Q 종료' $W) 'DarkGray'
    Write-Host ''
    Write-Seg (Format-Fit " $Msg" $W) 'Yellow'
    Write-Host ''
}

# ReadKey 기반의 간단한 한 줄 입력기 (메뉴와 같은 입력 방식을 써서 서로 충돌하지 않음)
# Enter = 확인, Esc = 취소($null 반환), Backspace = 지우기
function Read-TuiLine([string]$Prompt) {
    Write-Host $Prompt -NoNewline -ForegroundColor Cyan
    $buf = New-Object System.Text.StringBuilder
    while ($true) {
        $k = [Console]::ReadKey($true)
        switch ($k.Key) {
            'Enter'  { Write-Host ''; return $buf.ToString() }
            'Escape' { Write-Host ''; return $null }
            'Backspace' {
                if ($buf.Length -gt 0) {
                    $cw = Get-CharWidth $buf[$buf.Length - 1]
                    [void]$buf.Remove($buf.Length - 1, 1)
                    Write-Host (("`b" * $cw) + (' ' * $cw) + ("`b" * $cw)) -NoNewline
                }
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
    Write-Host "  $What 변경" -ForegroundColor Cyan
    Write-Host "  현재: $Current" -ForegroundColor DarkGray
    Write-Host ''
    Write-Host '  폴더 경로를 입력하거나 이 창으로 폴더를 끌어다 놓은 뒤 Enter 를 누르세요.'
    Write-Host '  Esc 를 누르거나 빈 줄로 Enter 를 누르면 취소됩니다.' -ForegroundColor DarkGray
    Write-Host ''
    try { $answer = Read-TuiLine '  > ' }
    finally { [Console]::CursorVisible = $false; Clear-Host }
    $answer
}

function Show-FullLog {
    Clear-Host
    $rows = [Math]::Max(5, [Console]::WindowHeight - 5)
    Write-Host ''
    Write-Host "  활동 기록  ($($script:LogPath))" -ForegroundColor Cyan
    Write-Host ''
    $lines = @(Read-DupLog $rows)
    if ($lines.Count -eq 0) { Write-Host '  아직 기록이 없습니다.' -ForegroundColor DarkGray }
    foreach ($l in $lines) { Write-Host "  $l" }
    Write-Host ''
    Write-Host '  아무 키나 누르면 메뉴로 돌아갑니다...' -ForegroundColor DarkGray
    [void][Console]::ReadKey($true)
    Clear-Host
}

function Invoke-MenuAction([int]$Index, $St) {
    switch ($Index) {
        0 { if ($St.State -eq 'ON') { Stop-Watcher } else { Start-Watcher } }
        1 {
            if ($St.Mode -eq 'Delete') { [void](Set-Setting 'mode' 'move'); '처리 방식: 중복 파일을 중복 폴더로 이동합니다.' }
            else { [void](Set-Setting 'mode' 'delete'); '처리 방식: 똑같은 중복 파일은 삭제합니다. (내용이 다르면 이동)' }
        }
        2 {
            if ($St.Keep -eq 'Newest') { [void](Set-Setting 'keep' 'oldest'); '보관 기준: 기존(오래된) 파일을 남깁니다.' }
            else { [void](Set-Setting 'keep' 'newest'); '보관 기준: 새로 받은(최근) 파일을 원래 이름으로 남깁니다.' }
        }
        3 {
            $a = Read-FolderInput '감시 폴더' $St.WatchFolder
            if ("$a".Trim()) { $c = Set-Setting 'watch' $a; "감시 폴더: $($c.WatchFolder)" } else { '취소했습니다.' }
        }
        4 {
            $a = Read-FolderInput '중복 폴더' $St.DuplicateFolder
            if ("$a".Trim()) { $c = Set-Setting 'dupdir' $a; "중복 폴더: $($c.DuplicateFolder)" } else { '취소했습니다.' }
        }
        5 { Set-Autostart ($St.Autostart -ne 'on') }
        6 { Show-FullLog; '' }
        7 {
            [void][System.IO.Directory]::CreateDirectory($St.DuplicateFolder)
            Invoke-Item -LiteralPath $St.DuplicateFolder
            '중복 폴더를 열었습니다.'
        }
    }
}

function Show-Menu {
    $sel = 0
    $msg = '환영합니다! 방향키와 Enter 로 조작하세요.'
    $count = 9
    $lastDraw = [datetime]::MinValue
    $lastWidth = -1
    $dirty = $true

    $oldTitle = $Host.UI.RawUI.WindowTitle
    $Host.UI.RawUI.WindowTitle = 'DupWatch'
    [Console]::CursorVisible = $false
    Clear-Host
    Initialize-Glyphs
    try {
        while ($true) {
            $winW = [Console]::WindowWidth
            $W = [Math]::Max(50, [Math]::Min(80, $winW - 1))
            if ($winW -ne $lastWidth) { Clear-Host; $lastWidth = $winW; $dirty = $true }

            # 키를 누른 뒤, 그리고 1초마다 다시 그려서 상태와 기록이 실시간으로 보이게 함
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
            if ($run -eq $count - 1) { break }   # 종료

            if ($run -eq 0) {   # 켜고 끄는 데는 몇 초 걸리므로 먼저 안내를 표시
                if ($st.State -eq 'ON') { $msg = '끄는 중...' } else { $msg = '켜는 중...' }
                Show-Screen $sel $msg $st $W
            }
            try { $msg = [string](Invoke-MenuAction $run $st | Select-Object -Last 1) }
            catch { $msg = "오류: $($_.Exception.Message)" }
            $lastWidth = -1   # 입력 화면 등에서 돌아온 뒤 화면을 깨끗이 다시 그림
        }
    } finally {
        [Console]::CursorVisible = $true
        $Host.UI.RawUI.WindowTitle = $oldTitle
        Clear-Host
        if (Test-Running) { Write-Host 'DupWatch 가 백그라운드에서 계속 실행 중입니다.' -ForegroundColor Green }
        else { Write-Host 'DupWatch 가 꺼져 있습니다.' -ForegroundColor DarkGray }
    }
}

# =====================================================================
#  시작 지점
# =====================================================================

try {
    switch ($Command) {
        'run'    { Invoke-WatcherLoop }
        'menu'   { Show-Menu }
        'start'  { Start-Watcher }
        'stop'   { Stop-Watcher }
        'status' {
            $s = Get-Status
            if ($s.State -eq 'ON') { $stText = '켜짐' } else { $stText = '꺼짐' }
            if ($s.Autostart -eq 'on') { $autoText = '켜짐' } else { $autoText = '꺼짐' }
            "상태:       $stText"
            "처리 방식:  $(Get-ModeText $s.Mode)"
            "보관 기준:  $(Get-KeepText $s.Keep)"
            "감시 폴더:  $($s.WatchFolder)"
            "중복 폴더:  $($s.DuplicateFolder)"
            "자동 시작:  $autoText"
            "기록 파일:  $($s.LogFile)"
        }
        'log'    { $l = Read-DupLog $Tail; if ($l) { $l } else { '아직 기록이 없습니다.' } }
        'autostart' {
            if ($Value -notin 'on', 'off') { throw '사용법: autostart on|off' }
            Set-Autostart ($Value -eq 'on')
        }
        { $_ -in 'mode', 'keep', 'watch', 'dupdir' } {
            if (-not $Value) { throw "사용법: $Command <값>" }
            $c = Set-Setting $Command $Value
            switch ($Command) {
                'mode'   { "처리 방식을 설정했습니다: $(Get-ModeText $c.Mode)" }
                'keep'   { "보관 기준을 설정했습니다: $(Get-KeepText $c.Keep)" }
                'watch'  { "감시 폴더를 설정했습니다: $($c.WatchFolder)" }
                'dupdir' { "중복 폴더를 설정했습니다: $($c.DuplicateFolder)" }
            }
        }
    }
} catch {
    Write-Host "오류: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}
