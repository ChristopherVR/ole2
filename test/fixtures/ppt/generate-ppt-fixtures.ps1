param([string]$OutputDirectory = $PSScriptRoot)
# Repository-owned synthetic content only; requires desktop PowerPoint and STA:
# pwsh -STA -NoProfile -File test/fixtures/ppt/generate-ppt-fixtures.ps1
# Office SaveAs embeds timestamps/version details; pin each generated revision by hash.
# SaveAs enum: https://learn.microsoft.com/en-us/office/vba/api/powerpoint.ppsaveasfiletype
# Metadata: https://learn.microsoft.com/en-us/office/vba/api/powerpoint.presentation.removedocumentinformation
$ErrorActionPreference = 'Stop'
$out = [IO.Path]::GetFullPath($OutputDirectory)
[void][IO.Directory]::CreateDirectory($out)
$app = $null
$deck = $null
try {
    $app = New-Object -ComObject PowerPoint.Application
    $app.AutomationSecurity = 3 # ForceDisable before creating any presentation.
    $app.DisplayAlerts = 1 # ppAlertsNone
    $deck = $app.Presentations.Add(0) # msoFalse: no visible document window.
    $deck.PageSetup.SlideWidth = 720
    $deck.PageSetup.SlideHeight = 405
    $slide = $deck.Slides.Add(1, 12) # ppLayoutBlank
    $title = $slide.Shapes.AddTextbox(1, 30, 25, 650, 50)
    $title.Name = 'FixtureTitle'
    $title.TextFrame.TextRange.Text = 'Native title fixture'
    $title.TextFrame.TextRange.Font.Name = 'Arial'
    $title.TextFrame.TextRange.Font.Size = 28
    $body = $slide.Shapes.AddTextbox(1, 30, 95, 650, 100)
    $body.Name = 'FixtureBody'
    $body.TextFrame.TextRange.Text = "Synthetic text body.`rLine two."
    $body.TextFrame.TextRange.Font.Name = 'Arial'
    $body.TextFrame.TextRange.Font.Size = 18
    $rectangle = $slide.Shapes.AddShape(1, 30, 230, 150, 80)
    $rectangle.Name = 'FixtureRectangle'
    $rectangle.Fill.ForeColor.RGB = 16711680
    $rectangle.Line.Visible = 0
    $slide.NotesPage.Shapes.Placeholders.Item(2).TextFrame.TextRange.Text = 'Synthetic notes retained.'
    $slide2 = $deck.Slides.Add(2, 12)
    $unicode = $slide2.Shapes.AddTextbox(1, 30, 25, 650, 100)
    $unicode.Name = 'FixtureUnicode'
    $unicode.TextFrame.TextRange.Text = 'Unicode Ω fixture'
    $unicode.TextFrame.TextRange.Font.Name = 'Arial'
    $unicode.TextFrame.TextRange.Font.Size = 24
    $wide = $slide2.Shapes.AddTextbox(1, 30, 150, 650, 100)
    $wide.Name = 'FixtureWideText'
    $wide.TextFrame.TextRange.Text = '日本語 café 😀'
    $wide.TextFrame.TextRange.Font.Name = 'Arial'
    $wide.TextFrame.TextRange.Font.Size = 18
    $slide2.NotesPage.Shapes.Placeholders.Item(2).TextFrame.TextRange.Text = 'Unicode notes retained.'
    # Strip inherited profile metadata. No personal documents are read.
    $deck.RemoveDocumentInformation(99) # ppRDIAll
    $props = $deck.BuiltinDocumentProperties
    foreach ($entry in @(@('Author', 'Fixture Author'), @('Title', 'Native binary PPT text fixture'))) {
        $prop = [System.__ComObject].InvokeMember('Item', 'GetProperty', $null, $props, @($entry[0]))
        [void][System.__ComObject].InvokeMember('Value', 'SetProperty', $null, $prop, @($entry[1]))
    }
    $deck.SaveAs((Join-Path $out 'native-text.ppt'), 1, 0) # ppSaveAsPresentation; no fonts embedded.
    Write-Output (Join-Path $out 'native-text.ppt')
}
finally {
    try {
        if ($null -ne $deck) { $deck.Close(); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($deck) }
    }
    finally {
        if ($null -ne $app) { $app.Quit(); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app) }
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}
