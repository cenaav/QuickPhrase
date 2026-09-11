<#
.SYNOPSIS
    Installs QuickPhrase.dotm into Word's STARTUP folder.

.DESCRIPTION
    Copies the template into %APPDATA%\Microsoft\Word\STARTUP so Word loads it
    at every start, and clears the "downloaded from the internet" mark that
    otherwise makes Word disable the macros silently.

    No admin rights needed - everything happens under the current user profile.

.PARAMETER Template
    Path to QuickPhrase.dotm. By default the script looks next to itself, then
    in ..\dist.
#>
[CmdletBinding()]
param(
    [string] $Template
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$appName = 'QuickPhrase'
$fileName = 'QuickPhrase.dotm'

# --- Locate the template ------------------------------------------------------

if (-not $Template) {
    $candidates = @(
        (Join-Path $PSScriptRoot $fileName)
        (Join-Path (Split-Path -Parent $PSScriptRoot) "dist\$fileName")
    )
    $Template = $candidates | Where-Object { Test-Path -LiteralPath $_ } | Select-Object -First 1
}

if (-not $Template -or -not (Test-Path -LiteralPath $Template)) {
    Write-Host ''
    Write-Host "Could not find $fileName." -ForegroundColor Red
    Write-Host ''
    Write-Host 'Put it next to this script, or download it from the Releases page:' -ForegroundColor Gray
    Write-Host '  https://github.com/cenaav/QuickPhrase/releases' -ForegroundColor Gray
    Write-Host ''
    Write-Host 'Building from source instead? Run build\build.ps1 first.' -ForegroundColor Gray
    exit 1
}

$Template = (Resolve-Path -LiteralPath $Template).Path

# --- Check Word is closed -----------------------------------------------------

if (Get-Process -Name WINWORD -ErrorAction SilentlyContinue) {
    Write-Host ''
    Write-Host 'Word is running. Close it before installing, then run this again.' -ForegroundColor Yellow
    Write-Host '(Word keeps STARTUP templates locked while it is open.)' -ForegroundColor Gray
    exit 1
}

# --- Copy ---------------------------------------------------------------------

$startup = Join-Path $env:APPDATA 'Microsoft\Word\STARTUP'
if (-not (Test-Path -LiteralPath $startup)) {
    New-Item -ItemType Directory -Path $startup -Force | Out-Null
}

$target = Join-Path $startup $fileName

Copy-Item -LiteralPath $Template -Destination $target -Force

# Word refuses to run macros from a file still carrying a mark-of-the-web
# alternate data stream, and does so without a useful message.
Unblock-File -LiteralPath $target

Write-Host ''
Write-Host "$appName installed." -ForegroundColor Green
Write-Host "  $target" -ForegroundColor Gray
Write-Host ''
Write-Host 'Start Word. A "QuickPhrase" tab appears next to Home.' -ForegroundColor Cyan
Write-Host ''
Write-Host 'If the tab is missing, macros are probably disabled:' -ForegroundColor Gray
Write-Host '  File > Options > Trust Center > Trust Center Settings > Macro Settings' -ForegroundColor Gray
Write-Host '  choose "Disable all macros with notification" (not "without notification")' -ForegroundColor Gray
Write-Host ''
Write-Host 'Your phrases are stored in:' -ForegroundColor Gray
Write-Host "  $env:APPDATA\$appName\snippets.json" -ForegroundColor Gray
