param([string]$OutputDirectory = $PSScriptRoot)
# Owned synthetic main-story field fixture, no source documents or external fields.
# Requires desktop Word: pwsh -STA -NoProfile -File this-script.ps1
# FileFormat0: Word97-2003 document. Header/footnote/endnote/table stories absent.
$ErrorActionPreference = 'Stop'
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
    $doc = $app.Documents.Add()
    $doc.Content.Text = "First paragraph plain text.`rSecond paragraph has a field: `rThird paragraph plain text."
    $doc.Content.Font.Name = 'Arial'
    $doc.Content.Font.Size = 12
    $range = $doc.Paragraphs.Item(2).Range.Duplicate
    $range.End = $range.End - 1
    $range.Collapse(0)
    # QUOTE with a fixed literal has no external reference or code execution.
    [void]$doc.Fields.Add($range, -1, 'QUOTE "Fixture field"', $false)
    $props = $doc.BuiltinDocumentProperties
    foreach ($entry in @(@('Author', 'Fixture Author'), @('Title', 'Synthetic main-story field fixture'))) {
        $prop = [System.__ComObject].InvokeMember('Item', 'GetProperty', $null, $props, @($entry[0]))
        [void][System.__ComObject].InvokeMember('Value', 'SetProperty', $null, $prop, @($entry[1]))
    }
    $doc.RemovePersonalInformation = $true
    $doc.EmbedTrueTypeFonts = $false
    $doc.SaveAs2((Join-Path $out 'main-field.doc'), 0, $false, '', $false)
    Write-Output (Join-Path $out 'main-field.doc')
}
finally {
    try {
        if ($null -ne $doc) { $doc.Close($false); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($doc) }
    }
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
