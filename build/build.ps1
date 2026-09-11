<#
.SYNOPSIS
    Builds dist/QuickPhrase.dotm from the text sources in src/.

.DESCRIPTION
    Runs in two stages:

      1. Word automation. Creates an empty macro-enabled template, imports the
         .bas modules, then builds the frmManager dialog control by control and
         attaches its code from src/forms/frmManager.code.vb. This stage is why
         the build needs Windows with Word installed - a vbaProject.bin cannot
         be produced any other way.

      2. Zip surgery. The saved .dotm is an OOXML package, so the customUI
         parts are added straight into it and _rels/.rels is rewritten to point
         at them. No Word involvement.

    Stage 1 requires "Trust access to the VBA project object model"
    (Word > File > Options > Trust Center > Trust Center Settings > Macro
    Settings). The script checks for it and explains how to turn it on.

.PARAMETER Output
    Destination .dotm. Defaults to dist/QuickPhrase.dotm.

.PARAMETER ShowWord
    Keep Word visible while building. Useful when a stage fails.

.PARAMETER EnableVbomTrust
    Set the AccessVBOM registry value that permits VBA project automation,
    instead of ticking it by hand in Word's Trust Center.

    This is the one setting the build needs. It applies to the current user
    only, needs no admin rights, and is reversible:

        reg add "HKCU\Software\Microsoft\Office\16.0\Word\Security" /v AccessVBOM /t REG_DWORD /d 0 /f

    Only code that edits macros is affected. The QuickPhrase add-in never needs
    it, so turning it back off after a successful build costs nothing.

.PARAMETER DisableVbomTrust
    Set AccessVBOM back to 0 and exit without building. Use it to restore the
    Trust Center setting once a build has succeeded.

.PARAMETER ExportVba
    Also copy the compiled VBA project out to package\word\vbaProject.bin.

    Commit that file, and every later release can be packed by build\pack.py on
    any OS - including a GitHub Actions Linux runner with no Word. Re-run this
    whenever the VBA source under src\ changes; ribbon XML and doc edits do
    not need a new blob.

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File build\build.ps1
#>
[CmdletBinding()]
param(
    [string] $Output,
    [switch] $ShowWord,
    [switch] $ExportVba,
    [switch] $EnableVbomTrust,
    [switch] $DisableVbomTrust
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# --- Paths --------------------------------------------------------------------

$repoRoot = Split-Path -Parent $PSScriptRoot
$srcRoot  = Join-Path $repoRoot 'src'

if (-not $Output) { $Output = Join-Path $repoRoot 'dist\QuickPhrase.dotm' }
$Output = [System.IO.Path]::GetFullPath($Output)

$modules = @(
    'JsonLite.bas'
    'UnicodeUI.bas'
    'SnippetStore.bas'
    'QuickPhraseRibbon.bas'
    'QuickPhraseMain.bas'
) | ForEach-Object { Join-Path $srcRoot "modules\$_" }

$formCode   = Join-Path $srcRoot 'forms\frmManager.code.vb'
$customUi14 = Join-Path $srcRoot 'customUI\customUI14.xml'
$customUi07 = Join-Path $srcRoot 'customUI\customUI.xml'

foreach ($f in @($modules + $formCode + $customUi14 + $customUi07)) {
    if (-not (Test-Path -LiteralPath $f)) { throw "Missing source file: $f" }
}

# --- Form layout --------------------------------------------------------------
# Ordered so tab order comes out sensibly. Coordinates are in points.

$formWidth  = 664
$formHeight = 398

$controls = @(
    @{ Type='Forms.Label.1';    Name='lblList';     Left=10;  Top=8;   Width=240; Height=14
       Props=@{ Caption='Phrases' } }

    @{ Type='Forms.ListBox.1';  Name='lstPhrases';  Left=10;  Top=24;  Width=240; Height=286
       Props=@{ IntegralHeight=$false } }

    @{ Type='Forms.CommandButton.1'; Name='btnNew';    Left=10;  Top=320; Width=58; Height=24
       Props=@{ Caption='New' } }
    @{ Type='Forms.CommandButton.1'; Name='btnDelete'; Left=72;  Top=320; Width=58; Height=24
       Props=@{ Caption='Delete' } }
    @{ Type='Forms.CommandButton.1'; Name='btnUp';     Left=134; Top=320; Width=52; Height=24
       Props=@{ Caption='Up' } }
    @{ Type='Forms.CommandButton.1'; Name='btnDown';   Left=190; Top=320; Width=60; Height=24
       Props=@{ Caption='Down' } }

    @{ Type='Forms.Label.1';    Name='lblLabel';    Left=264; Top=8;   Width=380; Height=14
       Props=@{ Caption='Button label (what you see on the ribbon)' } }
    @{ Type='Forms.TextBox.1';  Name='txtLabel';    Left=264; Top=24;  Width=380; Height=20 }

    @{ Type='Forms.Label.1';    Name='lblText';     Left=264; Top=52;  Width=380; Height=14
       Props=@{ Caption='Text inserted at the cursor' } }

    # Sized for a sentence or a short sign-off, which is what phrases actually
    # are. Longer text still scrolls.
    @{ Type='Forms.TextBox.1';  Name='txtText';     Left=264; Top=68;  Width=380; Height=150
       Props=@{ MultiLine=$true; WordWrap=$true; ScrollBars=2; EnterKeyBehavior=$true } }

    @{ Type='Forms.CheckBox.1'; Name='chkNewline';  Left=264; Top=226; Width=380; Height=18
       Props=@{ Caption='Start a new line after inserting this phrase' } }

    @{ Type='Forms.Label.1';    Name='lblSpacing';  Left=264; Top=252; Width=104; Height=14
       Props=@{ Caption='Add a space:' } }

    # fmStyleDropDownList (2) - pick from the list, no free typing.
    @{ Type='Forms.ComboBox.1'; Name='cboSpacing';  Left=370; Top=249; Width=150; Height=20
       Props=@{ Style=2 } }

    @{ Type='Forms.Label.1';    Name='lblHint';     Left=264; Top=278; Width=380; Height=32
       Props=@{ Caption='' } }

    @{ Type='Forms.CommandButton.1'; Name='btnSave';   Left=430; Top=320; Width=110; Height=24
       Props=@{ Caption='Save && Close'; Default=$true } }
    @{ Type='Forms.CommandButton.1'; Name='btnCancel'; Left=546; Top=320; Width=98;  Height=24
       Props=@{ Caption='Cancel'; Cancel=$true } }
)

# --- Stage 1: Word automation -------------------------------------------------

function Build-Template {
    param([string] $Destination)

    Write-Host 'Stage 1/2  Building VBA project with Word...' -ForegroundColor Cyan

    try {
        $word = New-Object -ComObject Word.Application
    } catch {
        throw "Could not start Word. Stage 1 needs Microsoft Word for Windows installed. ($($_.Exception.Message))"
    }

    $doc = $null
    try {
        $word.Visible = [bool] $ShowWord
        $word.DisplayAlerts = 0          # wdAlertsNone

        $doc = $word.Documents.Add()

        # wdFormatXMLTemplateMacroEnabled = 15
        $doc.SaveAs2($Destination, 15)

        # When access is denied, Word does not raise an error here - the property
        # just comes back empty. Both outcomes have to be handled, or the
        # failure surfaces later as a baffling "property 'Name' cannot be found".
        $project = $null
        try {
            $project = $doc.VBProject
        } catch {
            $project = $null
        }

        if ($null -eq $project) { throw (Get-VbomHelpText) }

        try {
            $components = $project.VBComponents
        } catch {
            throw (Get-VbomHelpText)
        }
        if ($null -eq $components) { throw (Get-VbomHelpText) }

        $project.Name = 'QuickPhrase'

        foreach ($module in $modules) {
            Write-Host "           import $(Split-Path -Leaf $module)"
            $project.VBComponents.Import($module) | Out-Null
        }

        Write-Host '           build frmManager'
        Add-ManagerForm -Project $project

        $doc.Save()
        Write-Host "           saved $Destination"
    }
    finally {
        if ($doc)  { try { $doc.Close(0) } catch { } }
        if ($word) { try { $word.Quit() }       catch { } }
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

function Add-ManagerForm {
    param($Project)

    # vbext_ct_MSForm = 3
    $form = $Project.VBComponents.Add(3)
    $form.Name = 'frmManager'
    $form.Properties('Caption').Value = 'QuickPhrase - Manage Phrases'
    $form.Properties('Width').Value  = $formWidth
    $form.Properties('Height').Value = $formHeight

    # Set the form font before adding controls so they all inherit it. Tahoma
    # renders Persian and Arabic correctly on every supported Windows version.
    $designer = $form.Designer
    $designer.Font.Name = 'Tahoma'
    $designer.Font.Size = 10

    foreach ($spec in $controls) {
        $control = $designer.Controls.Add($spec.Type, $spec.Name, $true)
        $control.Left   = $spec.Left
        $control.Top    = $spec.Top
        $control.Width  = $spec.Width
        $control.Height = $spec.Height

        if ($spec.ContainsKey('Props')) {
            foreach ($prop in $spec.Props.GetEnumerator()) {
                $control.($prop.Key) = $prop.Value
            }
        }
    }

    # The form starts with an empty code module, so the whole behaviour file
    # can go in as-is.
    $code = Get-Content -LiteralPath $formCode -Raw -Encoding UTF8
    $form.CodeModule.AddFromString($code)
}

# --- Stage 2: inject the ribbon parts -----------------------------------------

$relsTemplate = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
__EXISTING__
  <Relationship Id="rQuickPhraseUI2007" Type="http://schemas.microsoft.com/office/2006/relationships/ui/extensibility" Target="customUI/customUI.xml"/>
  <Relationship Id="rQuickPhraseUI2010" Type="http://schemas.microsoft.com/office/2007/relationships/ui/extensibility" Target="customUI/customUI14.xml"/>
</Relationships>
'@

function Add-CustomUi {
    param([string] $Package)

    Write-Host 'Stage 2/2  Injecting ribbon customisation...' -ForegroundColor Cyan

    Add-Type -AssemblyName System.IO.Compression.FileSystem
    Add-Type -AssemblyName System.IO.Compression

    $zip = [System.IO.Compression.ZipFile]::Open($Package, 'Update')
    try {
        # Read the existing package relationships, minus the closing tag, so
        # Word's own relationships survive.
        $relsEntry = $zip.GetEntry('_rels/.rels')
        if (-not $relsEntry) { throw '_rels/.rels not found - the .dotm looks corrupt.' }

        $reader = New-Object System.IO.StreamReader($relsEntry.Open())
        $existingRels = $reader.ReadToEnd()
        $reader.Dispose()

        $inner = [regex]::Match($existingRels, '(?s)<Relationships[^>]*>(.*)</Relationships>')
        if (-not $inner.Success) { throw 'Could not parse _rels/.rels.' }

        $newRels = $relsTemplate.Replace('__EXISTING__', $inner.Groups[1].Value.Trim())

        $relsEntry.Delete()
        Write-ZipText -Zip $zip -Path '_rels/.rels' -Text $newRels

        Write-ZipText -Zip $zip -Path 'customUI/customUI.xml' `
                      -Text (Get-Content -LiteralPath $customUi07 -Raw -Encoding UTF8)
        Write-ZipText -Zip $zip -Path 'customUI/customUI14.xml' `
                      -Text (Get-Content -LiteralPath $customUi14 -Raw -Encoding UTF8)

        Assert-XmlDefaultContentType -Zip $zip
    }
    finally {
        $zip.Dispose()
    }

    Write-Host '           customUI.xml + customUI14.xml added'
}

function Write-ZipText {
    param($Zip, [string] $Path, [string] $Text)

    $existing = $Zip.GetEntry($Path)
    if ($existing) { $existing.Delete() }

    $entry = $Zip.CreateEntry($Path)
    $stream = $entry.Open()
    try {
        # UTF8Encoding($false) keeps the BOM out; Word accepts either, but
        # plain UTF-8 is what every other OOXML part uses.
        $writer = New-Object System.IO.StreamWriter($stream, (New-Object System.Text.UTF8Encoding($false)))
        $writer.Write($Text)
        $writer.Flush()
        $writer.Dispose()
    }
    finally {
        $stream.Dispose()
    }
}

# The customUI parts rely on a Default entry for the xml extension. Word always
# writes one, but checking costs nothing and the failure would be baffling.
function Assert-XmlDefaultContentType {
    param($Zip)

    $entry = $Zip.GetEntry('[Content_Types].xml')
    if (-not $entry) { throw '[Content_Types].xml not found - the .dotm looks corrupt.' }

    $reader = New-Object System.IO.StreamReader($entry.Open())
    $types = $reader.ReadToEnd()
    $reader.Dispose()

    if ($types -notmatch 'Extension="xml"') {
        throw 'No Default content type for "xml" in the package; the ribbon parts would be ignored.'
    }
}

# --- Trust Center guidance ----------------------------------------------------

# Building the VBA project means automating the VBA editor, which Word gates
# behind a single Trust Center setting. Without it $doc.VBProject silently
# yields nothing, so these helpers are the difference between a five-second fix
# and a confusing hunt.

# Registry paths of every installed Word version, newest first.
function Get-WordSecurityKey {
    $keys = @()
    try {
        $keys = Get-ChildItem 'HKCU:\Software\Microsoft\Office' -ErrorAction SilentlyContinue |
                Where-Object { $_.PSChildName -match '^\d+\.\d+$' } |
                Sort-Object { [double] $_.PSChildName } -Descending
    } catch {
        return @()
    }

    $result = @()
    foreach ($key in $keys) {
        $result += [PSCustomObject]@{
            Version = $key.PSChildName
            Path    = "HKCU:\Software\Microsoft\Office\$($key.PSChildName)\Word\Security"
        }
    }
    return $result
}

# Reads AccessVBOM, or $null when the value has never been written. Under
# Set-StrictMode a missing property throws, hence the explicit check.
function Get-VbomValue {
    param([string] $SecurityPath)

    $props = Get-ItemProperty -Path $SecurityPath -ErrorAction SilentlyContinue
    if (-not $props) { return $null }
    if ($props.PSObject.Properties.Name -notcontains 'AccessVBOM') { return $null }
    return $props.AccessVBOM
}

function Enable-VbomTrust {
    $keys = Get-WordSecurityKey
    if ($keys.Count -eq 0) {
        throw 'No Word installation found under HKCU:\Software\Microsoft\Office.'
    }

    Write-Host 'Enabling VBA project access (current user only)...' -ForegroundColor Cyan

    foreach ($key in $keys) {
        if (-not (Test-Path -LiteralPath $key.Path)) {
            New-Item -Path $key.Path -Force | Out-Null
        }
        Set-ItemProperty -Path $key.Path -Name AccessVBOM -Value 1 -Type DWord
        Write-Host "           Office $($key.Version): AccessVBOM = 1"
    }

    Write-Host '           Turn it back off after the build with -DisableVbomTrust' -ForegroundColor Gray
}

function Disable-VbomTrust {
    foreach ($key in Get-WordSecurityKey) {
        if (Test-Path -LiteralPath $key.Path) {
            Set-ItemProperty -Path $key.Path -Name AccessVBOM -Value 0 -Type DWord
            Write-Host "Office $($key.Version): AccessVBOM = 0" -ForegroundColor Gray
        }
    }
}

function Get-VbomHelpText {
    $text = @'
Word blocked access to the VBA project object model.

This is the one setting the build needs. Either re-run with:

    build\build.ps1 -ExportVba -EnableVbomTrust

or tick it by hand:

    Word > File > Options > Trust Center > Trust Center Settings
         > Macro Settings > "Trust access to the VBA project object model"

then close Word and run this script again.

Only code that edits macros is affected. The QuickPhrase add-in never needs it,
so you can turn it back off once the build succeeds.

'@

    $found = $false
    foreach ($key in Get-WordSecurityKey) {
        $value = Get-VbomValue -SecurityPath $key.Path
        $found = $true
        if ($null -eq $value) {
            $text += "Office $($key.Version): AccessVBOM is not set (this is why it failed)`n"
        } elseif ($value -eq 1) {
            $text += "Office $($key.Version): AccessVBOM = 1 (enabled - is a Word window still open?)`n"
        } else {
            $text += "Office $($key.Version): AccessVBOM = $value (disabled)`n"
        }
    }

    if (-not $found) {
        $text += "No Word installation found under HKCU:\Software\Microsoft\Office.`n"
    }

    return $text
}

# --- Optional: export the compiled VBA project --------------------------------

# Lifts word/vbaProject.bin out of the freshly built template so it can be
# committed. That blob is the only part of the package Word alone can produce;
# with it in the repo, build\pack.py assembles releases anywhere.
function Export-VbaBlob {
    param([string] $Package)

    $blobPath = Join-Path $repoRoot 'package\word\vbaProject.bin'
    $blobDir = Split-Path -Parent $blobPath
    if (-not (Test-Path -LiteralPath $blobDir)) {
        New-Item -ItemType Directory -Path $blobDir -Force | Out-Null
    }

    Add-Type -AssemblyName System.IO.Compression.FileSystem

    $zip = [System.IO.Compression.ZipFile]::OpenRead($Package)
    try {
        $entry = $zip.GetEntry('word/vbaProject.bin')
        if (-not $entry) {
            throw 'word/vbaProject.bin is not in the built template; the VBA project did not compile.'
        }

        $source = $entry.Open()
        try {
            $target = [System.IO.File]::Create($blobPath)
            try { $source.CopyTo($target) } finally { $target.Dispose() }
        }
        finally { $source.Dispose() }
    }
    finally { $zip.Dispose() }

    $sizeKb = (Get-Item -LiteralPath $blobPath).Length / 1KB
    Write-Host ''
    Write-Host ("Exported VBA project: package\word\vbaProject.bin ({0:N1} KiB)" -f $sizeKb) -ForegroundColor Green
    Write-Host 'Commit it so CI can build releases without Word:' -ForegroundColor Gray
    Write-Host '  git add package/word/vbaProject.bin' -ForegroundColor Gray
}

# --- Main ---------------------------------------------------------------------

# Evaluate the version first: $IsWindows does not exist on PowerShell 5,
# and Set-StrictMode would trip over it.
if ($PSVersionTable.PSVersion.Major -ge 6 -and -not $IsWindows) {
    throw 'This build script needs Windows with Microsoft Word installed.'
}

if ($DisableVbomTrust) {
    Disable-VbomTrust
    exit 0
}

if ($EnableVbomTrust) { Enable-VbomTrust }

# Fail before starting Word rather than after. Launching Word takes seconds and
# the failure would be identical, just slower and with a stray process to clean
# up.
$trusted = $false
foreach ($key in Get-WordSecurityKey) {
    if ((Get-VbomValue -SecurityPath $key.Path) -eq 1) { $trusted = $true }
}
if (-not $trusted) {
    Write-Host ''
    Write-Host (Get-VbomHelpText) -ForegroundColor Yellow
    exit 1
}

if (Get-Process -Name WINWORD -ErrorAction SilentlyContinue) {
    Write-Host ''
    Write-Host 'Word is already running.' -ForegroundColor Yellow
    Write-Host 'Close every Word window and try again - an open instance can hold' -ForegroundColor Gray
    Write-Host 'the old Trust Center settings and lock the output file.' -ForegroundColor Gray
    exit 1
}

$outDir = Split-Path -Parent $Output
if (-not (Test-Path -LiteralPath $outDir)) {
    New-Item -ItemType Directory -Path $outDir -Force | Out-Null
}
if (Test-Path -LiteralPath $Output) { Remove-Item -LiteralPath $Output -Force }

# Errors here are almost always environmental - a Trust Center setting, a
# running Word, a locked file. A bare throw buries that advice under a
# PowerShell stack trace, so the message is printed on its own.
try {
    Build-Template -Destination $Output
    Add-CustomUi   -Package    $Output

    if ($ExportVba) { Export-VbaBlob -Package $Output }
}
catch {
    Write-Host ''
    Write-Host $_.Exception.Message -ForegroundColor Yellow
    Write-Host ''
    Write-Host "Re-run with -ShowWord to watch the build, or see docs/DEVELOPING.md." -ForegroundColor Gray
    exit 1
}

Write-Host ''
Write-Host "Built: $Output" -ForegroundColor Green
Write-Host 'Install it with install\Install.bat, or copy it into:' -ForegroundColor Gray
Write-Host "  $env:APPDATA\Microsoft\Word\STARTUP" -ForegroundColor Gray

if ($EnableVbomTrust) {
    Write-Host ''
    Write-Host 'Restore the Trust Center setting with:' -ForegroundColor Gray
    Write-Host '  build\build.ps1 -DisableVbomTrust' -ForegroundColor Gray
}
