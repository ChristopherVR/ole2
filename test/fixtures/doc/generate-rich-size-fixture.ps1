param(
    [string]$BaseFixturePath = (Join-Path $PSScriptRoot 'rich-runs.doc'),
    [string]$OutputDirectory = $PSScriptRoot
)
# Derive only from the pinned repository-owned synthetic fixture, never personal
# documents. Native Word97-2003 save materializes an exclusive18-point size run.
# Run with pwsh -STA. Source opens read-only, macros and link updates disabled.
$ErrorActionPreference = 'Stop'
$inputFile = (Resolve-Path -LiteralPath $BaseFixturePath).Path
if ((Get-FileHash -LiteralPath $inputFile -Algorithm SHA256).Hash.ToLowerInvariant() -ne '4b1cdf12f247fc235fdbee7232c3ed1d16d6380a98810b709b05e516cbfdd72f') {
    throw 'Expected the pinned owned rich-runs.doc fixture'
}
$out = [IO.Path]::GetFullPath($OutputDirectory)
[void][IO.Directory]::CreateDirectory($out)
$app = $null
$doc = $null
$oldLinks = $null
try {
    $app = New-Object -ComObject Word.Application
    $app.Visible = $false
    $app.DisplayAlerts = 0
    $app.AutomationSecurity = 3
    $oldLinks = $app.Options.UpdateLinksAtOpen
    $app.Options.UpdateLinksAtOpen = $false
    $doc = $app.Documents.Open($inputFile, $false, $true, $false, '', '', $false, '', '', 0, 0, $false, $false, 0, $true)
    $start = $doc.Paragraphs.Item(2).Range.Start
    if ($doc.Range($start, $start + 9).Text -ne 'Bold text') { throw 'Unexpected owned fixture text' }
    $doc.Range($start, $start + 9).Font.Size = 18
    $doc.RemovePersonalInformation = $true
    $doc.EmbedTrueTypeFonts = $false
    $doc.SaveAs2((Join-Path $out 'rich-size-runs.doc'), 0, $false, '', $false)
    Write-Output (Join-Path $out 'rich-size-runs.doc')
}
finally {
    try { if ($null -ne $doc) { $doc.Close($false); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($doc) } }
    finally {
        if ($null -ne $app) {
            if ($null -ne $oldLinks) { $app.Options.UpdateLinksAtOpen = $oldLinks }
            $app.Quit()
            [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)
        }
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}
