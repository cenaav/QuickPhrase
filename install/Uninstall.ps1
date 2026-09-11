<#
.SYNOPSIS
    Removes QuickPhrase.dotm from Word's STARTUP folder.

.DESCRIPTION
    Deletes the template only. Your phrase file in %APPDATA%\QuickPhrase is
    left alone, so reinstalling brings everything back. Pass -RemoveData to
    delete that too.
#>
[CmdletBinding()]
param(
    [switch] $RemoveData
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

if (Get-Process -Name WINWORD -ErrorAction SilentlyContinue) {
    Write-Host ''
    Write-Host 'Word is running. Close it before uninstalling, then run this again.' -ForegroundColor Yellow
    exit 1
}

$target = Join-Path $env:APPDATA 'Microsoft\Word\STARTUP\QuickPhrase.dotm'

if (Test-Path -LiteralPath $target) {
    Remove-Item -LiteralPath $target -Force
    Write-Host 'QuickPhrase removed from Word.' -ForegroundColor Green
} else {
    Write-Host 'QuickPhrase was not installed.' -ForegroundColor Gray
}

$dataDir = Join-Path $env:APPDATA 'QuickPhrase'

if ($RemoveData) {
    if (Test-Path -LiteralPath $dataDir) {
        Remove-Item -LiteralPath $dataDir -Recurse -Force
        Write-Host 'Phrase file deleted.' -ForegroundColor Green
    }
} elseif (Test-Path -LiteralPath $dataDir) {
    Write-Host ''
    Write-Host 'Your phrases were kept in:' -ForegroundColor Gray
    Write-Host "  $dataDir" -ForegroundColor Gray
    Write-Host 'Run this script with -RemoveData to delete them as well.' -ForegroundColor Gray
}
