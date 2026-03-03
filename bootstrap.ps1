$ErrorActionPreference = "Stop"

$okNoOp = 0x8A15002B

function Refresh-Path {
    $env:Path = [System.Environment]::GetEnvironmentVariable("Path", "Machine") +
        ";" + [System.Environment]::GetEnvironmentVariable("Path", "User")
}

function Install-WinGet {
    Write-Host "winget not found. Installing from GitHub..." -ForegroundColor Yellow

    $apiUrl = "https://api.github.com/repos/microsoft/winget-cli/releases/latest"
    $releaseJson = curl.exe -sL $apiUrl
    if ($LASTEXITCODE -ne 0) {
        throw "Failed to fetch winget release info from GitHub. Check your internet connection."
    }
    $release = $releaseJson | ConvertFrom-Json

    $msixBundleAsset = $release.assets | Where-Object { $_.name -match "\.msixbundle$" } | Select-Object -First 1
    if (-not $msixBundleAsset) {
        throw "Could not find msixbundle in the latest winget-cli release"
    }

    $downloadPath = "$env:TEMP\$($msixBundleAsset.name)"
    Write-Host "  Downloading $($msixBundleAsset.name)..." -ForegroundColor DarkGray
    curl.exe -sL -o $downloadPath $msixBundleAsset.browser_download_url
    if ($LASTEXITCODE -ne 0 -or -not (Test-Path $downloadPath)) {
        throw "Failed to download winget installer. Check your internet connection."
    }

    Write-Host "  Installing winget..." -ForegroundColor DarkGray
    try {
        Add-AppxPackage -Path $downloadPath
    }
    catch {
        throw "Failed to install winget: $_`nYou may need to manually install dependencies (e.g. Microsoft.VCLibs) or run as administrator."
    }

    Remove-Item $downloadPath -ErrorAction SilentlyContinue

    Refresh-Path

    if (-not (Get-Command winget -ErrorAction SilentlyContinue)) {
        throw "winget still not available after installation"
    }

    Write-Host "  winget installed successfully." -ForegroundColor Green
}

if (-not (Get-Command winget -ErrorAction SilentlyContinue)) {
    Install-WinGet
}

function Install-WithWinGet {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Id
    )

    winget install --id $Id --exact --silent --disable-interactivity --accept-package-agreements --accept-source-agreements --scope user
    if ($LASTEXITCODE -ne 0 -and $LASTEXITCODE -ne $okNoOp) {
        throw "winget install failed for '$Id' with exit code $LASTEXITCODE"
    }
}

Install-WithWinGet -Id "Git.Git"
Install-WithWinGet -Id "twpayne.chezmoi"

Refresh-Path

$chezmoiCmd = Get-Command chezmoi -ErrorAction SilentlyContinue
if (-not $chezmoiCmd) {
    throw "chezmoi.exe not found after installation"
}

& $chezmoiCmd.Source init --apply https://github.com/Dreaming-Codes/windowsdots.git
