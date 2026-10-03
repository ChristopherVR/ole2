param([string]$OutputDirectory = $PSScriptRoot)
# Repository-owned synthetic content only; Word 97-2003 FileFormat 0.
# Run with pwsh -STA. No fields, external references, objects, or source documents.
$ErrorActionPreference = 'Stop'
$out = [IO.Path]::GetFullPath($OutputDirectory)
[void][IO.Directory]::CreateDirectory($out)
$app = $null
$doc = $null
try {
    $app = New-Object -ComObject Word.Application
    $app.Visible = $false
    $app.DisplayAlerts = 0
    $app.AutomationSecurity = 3
    $doc = $app.Documents.Add()
    $doc.Styles.Item(-1).Font.Name = 'Arial'
    $doc.Styles.Item(-1).Font.Size = 12
    $doc.Content.Text = "First plain paragraph.`rBold text, italic text, plain text.`rLast plain paragraph."
    $doc.Content.Style = $doc.Styles.Item(-1)
    $doc.Content.Font.Name = 'Arial'
    $doc.Content.Font.Size = 12
    $doc.Content.Font.Bold = 0
    $doc.Content.Font.Italic = 0
    $start = $doc.Paragraphs.Item(2).Range.Start
    $doc.Range($start, $start + 9).Font.Bold = -1
    $doc.Range($start + 11, $start + 22).Font.Italic = -1
    $props = $doc.BuiltinDocumentProperties
    foreach ($entry in @(@('Author', 'Fixture Author'), @('Title', 'Synthetic direct formatting fixture'))) {
        $prop = [System.__ComObject].InvokeMember('Item', 'GetProperty', $null, $props, @($entry[0]))
        [void][System.__ComObject].InvokeMember('Value', 'SetProperty', $null, $prop, @($entry[1]))
    }
    $doc.RemovePersonalInformation = $true
    $doc.EmbedTrueTypeFonts = $false
    $doc.SaveAs2((Join-Path $out 'rich-runs.doc'), 0, $false, '', $false)
    Write-Output (Join-Path $out 'rich-runs.doc')
}
finally {
    try { if ($null -ne $doc) { $doc.Close($false); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($doc) } }
    finally {
        if ($null -ne $app) { $app.Quit(); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app) }
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}
