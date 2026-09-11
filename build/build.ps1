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

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File build\build.ps1
#>
[CmdletBinding()]
param(
    [string] $Output,
    [switch] $ShowWord
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
$formHeight = 430

$controls = @(
    @{ Type='Forms.Label.1';    Name='lblList';     Left=10;  Top=8;   Width=240; Height=14
       Props=@{ Caption='Phrases' } }

    @{ Type='Forms.ListBox.1';  Name='lstPhrases';  Left=10;  Top=24;  Width=240; Height=320
       Props=@{ IntegralHeight=$false } }

    @{ Type='Forms.CommandButton.1'; Name='btnNew';    Left=10;  Top=352; Width=58; Height=24
       Props=@{ Caption='New' } }
    @{ Type='Forms.CommandButton.1'; Name='btnDelete'; Left=72;  Top=352; Width=58; Height=24
       Props=@{ Caption='Delete' } }
    @{ Type='Forms.CommandButton.1'; Name='btnUp';     Left=134; Top=352; Width=52; Height=24
       Props=@{ Caption='Up' } }
    @{ Type='Forms.CommandButton.1'; Name='btnDown';   Left=190; Top=352; Width=60; Height=24
       Props=@{ Caption='Down' } }

    @{ Type='Forms.Label.1';    Name='lblLabel';    Left=264; Top=8;   Width=380; Height=14
       Props=@{ Caption='Button label (what you see on the ribbon)' } }
    @{ Type='Forms.TextBox.1';  Name='txtLabel';    Left=264; Top=24;  Width=380; Height=20 }

    @{ Type='Forms.Label.1';    Name='lblText';     Left=264; Top=52;  Width=380; Height=14
       Props=@{ Caption='Text inserted at the cursor' } }
    @{ Type='Forms.TextBox.1';  Name='txtText';     Left=264; Top=68;  Width=380; Height=230
       Props=@{ MultiLine=$true; WordWrap=$true; ScrollBars=2; EnterKeyBehavior=$true } }

    @{ Type='Forms.Label.1';    Name='lblHint';     Left=264; Top=304; Width=380; Height=42
       Props=@{ Caption='' } }

    @{ Type='Forms.CommandButton.1'; Name='btnSave';   Left=430; Top=352; Width=110; Height=24
       Props=@{ Caption='Save && Close'; Default=$true } }
    @{ Type='Forms.CommandButton.1'; Name='btnCancel'; Left=546; Top=352; Width=98;  Height=24
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

        try {
            $project = $doc.VBProject
        } catch {
            throw @'
Word blocked access to the VBA project.

Turn it on once:
  Word > File > Options > Trust Center > Trust Center Settings
       > Macro Settings > tick "Trust access to the VBA project object model"

Then run this script again. You can untick it afterwards.
'@
        }

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

# --- Main ---------------------------------------------------------------------

# Evaluate the version first: $IsWindows does not exist on PowerShell 5,
# and Set-StrictMode would trip over it.
if ($PSVersionTable.PSVersion.Major -ge 6 -and -not $IsWindows) {
    throw 'This build script needs Windows with Microsoft Word installed.'
}

$outDir = Split-Path -Parent $Output
if (-not (Test-Path -LiteralPath $outDir)) {
    New-Item -ItemType Directory -Path $outDir -Force | Out-Null
}
if (Test-Path -LiteralPath $Output) { Remove-Item -LiteralPath $Output -Force }

Build-Template -Destination $Output
Add-CustomUi   -Package    $Output

Write-Host ''
Write-Host "Built: $Output" -ForegroundColor Green
Write-Host 'Install it with install\Install.bat, or copy it into:' -ForegroundColor Gray
Write-Host "  $env:APPDATA\Microsoft\Word\STARTUP" -ForegroundColor Gray
