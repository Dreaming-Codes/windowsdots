$ErrorActionPreference = "Stop"

$okNoOp = 0x8A15002B

function Refresh-Path {
    $env:Path = [System.Environment]::GetEnvironmentVariable("Path", "Machine") +
        ";" + [System.Environment]::GetEnvironmentVariable("Path", "User")
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

& $chezmoiCmd.Source init --apply git@github.com:Dreaming-Codes/windowsdots
