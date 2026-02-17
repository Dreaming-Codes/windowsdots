$ErrorActionPreference = "Stop"

$winget = (Get-Command winget).Source
$espansoDir = "$env:LOCALAPPDATA\Programs\Espanso"
$espansoExe = "$espansoDir\espanso.cmd"

function Refresh-Path {
    $env:Path = [System.Environment]::GetEnvironmentVariable("Path", "Machine") +
        ";" + [System.Environment]::GetEnvironmentVariable("Path", "User")
}

# =============================================================
# Phase 1: Blocking installs
# =============================================================

Write-Host "`n[1/6] Installing Git..." -ForegroundColor Cyan
& $winget install --id Git.Git -e --scope user --accept-source-agreements --accept-package-agreements -h

Write-Host "`n[2/6] Installing Espanso..." -ForegroundColor Cyan
& $winget install --id Espanso.Espanso -e --scope user --accept-source-agreements --accept-package-agreements -h
Refresh-Path

Write-Host "`n[3/6] Installing Helium Browser..." -ForegroundColor Cyan
& $winget install --id ImputNet.Helium -e --scope user --accept-source-agreements --accept-package-agreements -h
Write-Host "  Helium installed." -ForegroundColor Green

# =============================================================
# Phase 2: Set Helium as default browser
# =============================================================

Write-Host "`n[4/6] Setting Helium as default browser..." -ForegroundColor Cyan

$setFtaDir = "$env:TEMP\SetUserFTA"
$setFtaExe = "$setFtaDir\SetUserFTA.exe"

if (-not (Test-Path $setFtaExe)) {
    $zipPath = "$env:TEMP\SetUserFTA.zip"
    curl.exe -sL -o $zipPath "https://setuserfta.com/SetUserFTA.zip"
    Expand-Archive -Path $zipPath -DestinationPath $setFtaDir -Force
    Remove-Item $zipPath
}

# Background job to auto-dismiss SetUserFTA free version popup
$dismissJob = Start-Job -ScriptBlock {
    Add-Type -AssemblyName System.Windows.Forms
    while ($true) {
        $wshell = New-Object -ComObject WScript.Shell
        $popup = Get-Process | Where-Object { $_.MainWindowTitle -match "SetUserFTA" }
        foreach ($p in $popup) {
            if ($p.MainWindowHandle -ne [IntPtr]::Zero) {
                $wshell.AppActivate($p.Id) | Out-Null
                Start-Sleep -Milliseconds 100
                [System.Windows.Forms.SendKeys]::SendWait("{ENTER}")
            }
        }
        Start-Sleep -Milliseconds 500
    }
}

# Detect Helium ProgID
$heliumProgId = $null
@(
    "Registry::HKEY_CURRENT_USER\Software\Classes",
    "Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Classes"
) | ForEach-Object {
    if (-not $heliumProgId) {
        $match = Get-ChildItem $_ -ErrorAction SilentlyContinue |
            Where-Object { $_.PSChildName -match "^HeliumHTML|^HeliumHTM" } |
            Select-Object -First 1
        if ($match) { $heliumProgId = $match.PSChildName }
    }
}

if (-not $heliumProgId) {
    Write-Warning "Helium ProgID not found, falling back to 'HeliumHTML'."
    $heliumProgId = "HeliumHTML"
}

Write-Host "  ProgID: $heliumProgId" -ForegroundColor DarkGray

$extensions = @(".htm", ".html", ".xhtml", ".shtml", ".svg", ".webp", ".pdf")
$protocols = @("http", "https", "ftp")

foreach ($ext in $extensions) {
    & $setFtaExe $heliumProgId $ext
    Write-Host "  Set $ext" -ForegroundColor DarkGray
}
foreach ($proto in $protocols) {
    & $setFtaExe $heliumProgId $proto
    Write-Host "  Set $proto" -ForegroundColor DarkGray
}

# Stop the dismiss job
Stop-Job $dismissJob -ErrorAction SilentlyContinue
Remove-Job $dismissJob -Force -ErrorAction SilentlyContinue

$signature = @'
[DllImport("shell32.dll")]
public static extern void SHChangeNotify(int wEventId, int uFlags, IntPtr dwItem1, IntPtr dwItem2);
'@
$shell32 = Add-Type -MemberDefinition $signature -Name "Shell32" -Namespace "Win32" -PassThru
$shell32::SHChangeNotify(0x08000000, 0, [IntPtr]::Zero, [IntPtr]::Zero)

Write-Host "  Helium set as default ($($extensions.Count) extensions, $($protocols.Count) protocols)." -ForegroundColor Green

# =============================================================
# Phase 3: Espanso templates + start espanso
# =============================================================

Start-Process -FilePath $espansoExe -ArgumentList "start" -WindowStyle Hidden
Write-Host "  Espanso started." -ForegroundColor Green

# =============================================================
# Phase 4: Secondary installs
# =============================================================

Write-Host "`n[5/6] Installing Zed Editor..." -ForegroundColor Cyan
& $winget install --id ZedIndustries.Zed -e --accept-source-agreements --accept-package-agreements -h
Write-Host "  Zed installed." -ForegroundColor Green

Write-Host "  Configuring Zed..." -ForegroundColor Cyan
$zedConfigDir = "$env:APPDATA\Zed"
if (-not (Test-Path $zedConfigDir)) {
    New-Item -ItemType Directory -Path $zedConfigDir -Force | Out-Null
}

$content = @'
// Zed settings
//
// For information on how to configure Zed, see the Zed
// documentation: https://zed.dev/docs/configuring-zed
//
// To see all of Zed's default settings without changing your
// custom settings, run `zed: open default settings` from the
// command palette (cmd-shift-p / ctrl-shift-p)
{
  "helix_mode": true,
  "session": {
    "trust_all_worktrees": true
  },
  "icon_theme": "Zed (Default)",
  "ui_font_size": 16,
  "buffer_font_size": 15,
  "theme": {
    "mode": "dark",
    "light": "One Light",
    "dark": "One Dark"
  }
}
'@

[System.IO.File]::WriteAllText("$zedConfigDir\settings.json", $content, [System.Text.UTF8Encoding]::new($false))
Write-Host "  Zed configured." -ForegroundColor Green

Refresh-Path

# =============================================================
# Phase 5: Windows personalization & taskbar
# =============================================================

Write-Host "`n[6/6] Configuring Windows appearance and taskbar..." -ForegroundColor Cyan

# --- Dark mode ---
$themePath = "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize"
Set-ItemProperty -Path $themePath -Name "AppsUseLightTheme" -Value 0 -Type DWord
Set-ItemProperty -Path $themePath -Name "SystemUsesLightTheme" -Value 0 -Type DWord
Write-Host "  Dark mode enabled." -ForegroundColor DarkGray

# --- Accent color: #E81123 ---
Set-ItemProperty -Path "HKCU:\Control Panel\Desktop" -Name "AutoColorization" -Value 0 -Type DWord

# Convert to signed Int32 safely (0xFF2311E8 overflows [int])
$abgr = [BitConverter]::ToInt32([byte[]](0xE8, 0x11, 0x23, 0xFF), 0)

$dwmPath = "HKCU:\SOFTWARE\Microsoft\Windows\DWM"
Set-ItemProperty -Path $dwmPath -Name "AccentColor" -Value $abgr -Type DWord
Set-ItemProperty -Path $dwmPath -Name "ColorizationColor" -Value $abgr -Type DWord
Set-ItemProperty -Path $dwmPath -Name "ColorizationAfterglow" -Value $abgr -Type DWord
# Show accent on title bars and taskbar
Set-ItemProperty -Path $dwmPath -Name "ColorPrevalence" -Value 1 -Type DWord

$accentPath = "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Accent"
Set-ItemProperty -Path $accentPath -Name "AccentColorMenu" -Value $abgr -Type DWord
Set-ItemProperty -Path $accentPath -Name "StartColorMenu" -Value $abgr -Type DWord

# AccentPalette: 8 RGBA entries (32 bytes) - gradient from light to dark
# Built around #E81123
$accentPalette = [byte[]](
    0xFB, 0xCB, 0xCE, 0x00,  # lightest
    0xF4, 0x97, 0x9D, 0x00,
    0xEE, 0x5F, 0x67, 0x00,
    0xE8, 0x11, 0x23, 0x00,  # base
    0xBF, 0x0E, 0x1D, 0x00,
    0x96, 0x0B, 0x17, 0x00,
    0x6E, 0x08, 0x11, 0x00,
    0x4C, 0x06, 0x0C, 0x00   # darkest
)
Set-ItemProperty -Path $accentPath -Name "AccentPalette" -Value $accentPalette -Type Binary

# Show accent on Start and taskbar
$themePath = "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize"
Set-ItemProperty -Path $themePath -Name "ColorPrevalence" -Value 1 -Type DWord

# Broadcast setting change so running apps pick it up
Add-Type -Namespace Win32 -Name NativeMethods -MemberDefinition @'
[DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
public static extern IntPtr SendMessageTimeout(
    IntPtr hWnd, uint Msg, UIntPtr wParam, string lParam,
    uint fuFlags, uint uTimeout, out UIntPtr lpdwResult);
'@
$HWND_BROADCAST = [IntPtr]0xFFFF
$WM_SETTINGCHANGE = 0x001A
$result = [UIntPtr]::Zero
[Win32.NativeMethods]::SendMessageTimeout(
    $HWND_BROADCAST, $WM_SETTINGCHANGE, [UIntPtr]::Zero,
    "ImmersiveColorSet", 0x0002, 5000, [ref]$result
) | Out-Null

Write-Host "  Accent color set to #E81123." -ForegroundColor DarkGray

# --- Left-aligned taskbar ---
$advPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"
Set-ItemProperty -Path $advPath -Name "TaskbarAl" -Value 0 -Type DWord
Write-Host "  Taskbar left-aligned." -ForegroundColor DarkGray

# --- Hide search bar ---
$searchPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Search"
if (-not (Test-Path $searchPath)) {
    New-Item -Path $searchPath -Force | Out-Null
}
Set-ItemProperty -Path $searchPath -Name "SearchboxTaskbarMode" -Value 0 -Type DWord
Write-Host "  Search bar hidden." -ForegroundColor DarkGray

# --- Hide other taskbar clutter (Copilot, Widgets, Task View, Chat) ---
Set-ItemProperty -Path $advPath -Name "ShowCopilotButton" -Value 0 -Type DWord
Set-ItemProperty -Path $advPath -Name "ShowTaskViewButton" -Value 0 -Type DWord
Set-ItemProperty -Path $advPath -Name "TaskbarMn" -Value 0 -Type DWord
Write-Host "  Copilot/Widgets/TaskView/Chat hidden." -ForegroundColor DarkGray

# --- Auto-hide taskbar ---
# Reliable method: set StuckRects3 "Settings" byte 8 to 3 (auto-hide on)
$stuckRectsPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\StuckRects3"
if (Test-Path $stuckRectsPath) {
    $settings = (Get-ItemProperty -Path $stuckRectsPath -Name "Settings" -ErrorAction Stop).Settings
    if ($settings -and $settings.Length -gt 8) {
        $settings[8] = 3
        Set-ItemProperty -Path $stuckRectsPath -Name "Settings" -Value $settings -Type Binary
    }
}
Write-Host "  Taskbar auto-hide enabled." -ForegroundColor DarkGray

# --- Taskbar pinned apps ---
Write-Host "  Configuring taskbar pins..." -ForegroundColor DarkGray

$taskbarDir = "$env:APPDATA\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar"
if (-not (Test-Path $taskbarDir)) {
    New-Item -ItemType Directory -Path $taskbarDir -Force | Out-Null
}

# Clear existing pins
Remove-Item "$taskbarDir\*" -Force -ErrorAction SilentlyContinue

# Clear the binary taskbar pin state in registry
$taskbandPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Taskband"
if (Test-Path $taskbandPath) {
    Remove-ItemProperty -Path $taskbandPath -Name "Favorites" -ErrorAction SilentlyContinue
    Remove-ItemProperty -Path $taskbandPath -Name "FavoritesResolve" -ErrorAction SilentlyContinue
    Remove-ItemProperty -Path $taskbandPath -Name "FavoritesVersion" -ErrorAction SilentlyContinue
    Remove-ItemProperty -Path $taskbandPath -Name "FavoritesChanges" -ErrorAction SilentlyContinue
}

# Restart Explorer to apply all changes
Write-Host "  Restarting Explorer..." -ForegroundColor DarkGray
Stop-Process -Name explorer -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 2
Start-Process explorer

Write-Host "  Windows personalization applied." -ForegroundColor Green

# =============================================================
# Phase 7: Install whkd, komorebi, and masir
# =============================================================

Write-Host "`n[8/8] Installing whkd, komorebi, and masir..." -ForegroundColor Cyan

$toolsDir = "$env:LOCALAPPDATA\Tools"
$whkdDir = "$toolsDir\whkd"
$komorebiDir = "$toolsDir\komorebi"
$masirDir = "$toolsDir\masir"
$whkdUrl = "https://github.com/LGUG2Z/whkd/releases/download/v0.2.10/whkd-0.2.10-x86_64-pc-windows-msvc.zip"
$komorebiLatest = (curl.exe -sL "https://api.github.com/repos/LGUG2Z/komorebi/releases/latest" | ConvertFrom-Json).tag_name
$komorebiVersion = $komorebiLatest.TrimStart("v")
$komorebiUrl = "https://github.com/LGUG2Z/komorebi/releases/download/$komorebiLatest/komorebi-$komorebiVersion-x86_64-pc-windows-msvc.zip"
$masirLatest = (curl.exe -sL "https://api.github.com/repos/LGUG2Z/masir/releases/latest" | ConvertFrom-Json).tag_name
$masirVersion = $masirLatest.TrimStart("v")
$masirUrl = "https://github.com/LGUG2Z/masir/releases/download/$masirLatest/masir-$masirVersion-x86_64-pc-windows-msvc.zip"

# Create tools directory
if (-not (Test-Path $toolsDir)) {
    New-Item -ItemType Directory -Path $toolsDir -Force | Out-Null
}

# Download and install whkd
if (Test-Path (Join-Path $whkdDir "whkd.exe")) {
    Write-Host "  whkd already installed. Skipping download." -ForegroundColor DarkGray
} else {
    Write-Host "  Downloading whkd..." -ForegroundColor DarkGray
    $whkdZip = "$env:TEMP\whkd.zip"
    curl.exe -sL -o $whkdZip $whkdUrl

    if (Test-Path $whkdDir) {
        Remove-Item -Recurse -Force $whkdDir
    }
    Expand-Archive -Path $whkdZip -DestinationPath $whkdDir -Force
    Remove-Item $whkdZip
}

# Download and install komorebi
if (Test-Path (Join-Path $komorebiDir "komorebic.exe")) {
    Write-Host "  komorebi already installed. Skipping download." -ForegroundColor DarkGray
} else {
    Write-Host "  Downloading komorebi..." -ForegroundColor DarkGray
    $komorebiZip = "$env:TEMP\komorebi.zip"
    curl.exe -sL -o $komorebiZip $komorebiUrl

    if (Test-Path $komorebiDir) {
        Remove-Item -Recurse -Force $komorebiDir
    }
    Expand-Archive -Path $komorebiZip -DestinationPath $komorebiDir -Force
    Remove-Item $komorebiZip
}

# Download and install masir
if (Test-Path (Join-Path $masirDir "masir.exe")) {
    Write-Host "  masir already installed. Skipping download." -ForegroundColor DarkGray
} else {
    Write-Host "  Downloading masir..." -ForegroundColor DarkGray
    $masirZip = "$env:TEMP\masir.zip"
    curl.exe -sL -o $masirZip $masirUrl

    if (Test-Path $masirDir) {
        Remove-Item -Recurse -Force $masirDir
    }
    Expand-Archive -Path $masirZip -DestinationPath $masirDir -Force
    Remove-Item $masirZip
}

# Add to PATH
Write-Host "  Adding to PATH..." -ForegroundColor DarkGray
$userPath = [System.Environment]::GetEnvironmentVariable("Path", "User")
$pathEntries = $userPath -split ";" | Where-Object { $_ -ne "" }

# Remove old entries if they exist
$pathEntries = $pathEntries | Where-Object { $_ -notmatch "whkd|komorebi|masir" }

# Add new entries
$pathEntries += $whkdDir
$pathEntries += $komorebiDir
$pathEntries += $masirDir

$newPath = ($pathEntries -join ";").TrimEnd(";")
[System.Environment]::SetEnvironmentVariable("Path", $newPath, "User")
Refresh-Path

$komorebicCmd = Join-Path $komorebiDir "komorebic.exe"
if (Test-Path $komorebicCmd) {
    $autostart = Start-Process -FilePath $komorebicCmd -ArgumentList "enable-autostart --whkd --bar --masir" -Wait -NoNewWindow -PassThru
    if ($autostart.ExitCode -eq 0) {
        Write-Host "  Komorebi autostart enabled." -ForegroundColor DarkGray
    } else {
        Write-Warning "Komorebi autostart may already be enabled (exit code: $($autostart.ExitCode)). Continuing."
    }
} else {
    Write-Warning "komorebic.exe not found; skipping autostart setup."
}

# Start komorebi if not already running
if (-not (Get-Process -Name "komorebi" -ErrorAction SilentlyContinue)) {
    Write-Host "  Starting komorebi..." -ForegroundColor DarkGray
    Start-Process -FilePath $komorebicCmd -ArgumentList "start --bar --whkd --masir" -NoNewWindow
} else {
    Write-Host "  komorebi is already running. Skipping start." -ForegroundColor DarkGray
}
