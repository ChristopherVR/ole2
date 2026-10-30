param([string]$OutputDirectory = $PSScriptRoot)
# Repository-owned synthetic content only; Word 97-2003 (FileFormat 0) via Word COM.
# Run with pwsh -STA. Produces formatting-rich.doc, formatting-styles.doc,
# formatting-lists.doc and formatting-unicode.doc. No fields, links, objects or source documents.
$ErrorActionPreference = 'Stop'
$out = [IO.Path]::GetFullPath($OutputDirectory)
[void][IO.Directory]::CreateDirectory($out)
function Rgb([int]$r, [int]$g, [int]$b) { return $r + ($g * 256) + ($b * 65536) }
$app = New-Object -ComObject Word.Application
$app.Visible = $false
$app.DisplayAlerts = 0
$app.AutomationSecurity = 3
function Save([object]$doc, [string]$name) {
    $doc.RemovePersonalInformation = $true
    $doc.EmbedTrueTypeFonts = $false
    $doc.SaveAs2((Join-Path $out $name), 0, $false, '', $false)
    $doc.Close($false)
}
try {
    # ---- formatting-rich.doc: direct character and paragraph formatting ----
    $doc = $app.Documents.Add()
    $doc.Content.Text = "Centered title`rbold italic underline strike dstrike caps small hidden super sub colour highlight courier big spaced`rIndented justified paragraph with borders shading and tabs.`rHanging indent paragraph."
    $doc.Content.Font.Name = 'Arial'
    $doc.Content.Font.Size = 12
    $doc.DefaultTabStop = 40.5
    $p1 = $doc.Paragraphs.Item(1)
    $p1.Alignment = 1
    $p1.SpaceBefore = 12
    $p1.SpaceAfter = 6
    $p1.LineSpacingRule = 5
    $p1.LineSpacing = 18
    $p2 = $doc.Paragraphs.Item(2).Range
    $words = @('bold','italic','underline','strike','dstrike','caps','small','hidden','super','sub','colour','highlight','courier','big','spaced')
    $pos = $p2.Start
    $r = @{}
    foreach ($w in $words) { $r[$w] = $doc.Range($pos, $pos + $w.Length); $pos += $w.Length + 1 }
    $r['bold'].Font.Bold = -1
    $r['italic'].Font.Italic = -1
    $r['underline'].Font.Underline = 3
    $r['strike'].Font.StrikeThrough = -1
    $r['dstrike'].Font.DoubleStrikeThrough = -1
    $r['caps'].Font.AllCaps = -1
    $r['small'].Font.SmallCaps = -1
    $r['hidden'].Font.Hidden = -1
    $r['super'].Font.Superscript = -1
    $r['sub'].Font.Subscript = -1
    $r['colour'].Font.Color = (Rgb 200 16 32)
    $r['highlight'].HighlightColorIndex = 7
    $r['courier'].Font.Name = 'Courier New'
    $r['big'].Font.Size = 20
    $r['spaced'].Font.Spacing = 3
    $p3 = $doc.Paragraphs.Item(3)
    $p3.Alignment = 3
    $p3.LeftIndent = 36
    $p3.RightIndent = 18
    $p3.FirstLineIndent = 18
    $p3.KeepWithNext = -1
    $p3.KeepTogether = -1
    $p3.PageBreakBefore = -1
    $p3.WidowControl = 0
    $p3.OutlineLevel = 2
    $p3.Borders.Item(-1).LineStyle = 1
    $p3.Borders.Item(-1).LineWidth = 8
    $p3.Borders.Item(-3).LineStyle = 1
    $p3.Shading.BackgroundPatternColor = (Rgb 255 255 0)
    [void]$p3.TabStops.Add(72, 1, 1)
    [void]$p3.TabStops.Add(216, 2, 0)
    $p4 = $doc.Paragraphs.Item(4)
    $p4.LeftIndent = 54
    $p4.FirstLineIndent = -18
    $p4.Alignment = 2
    Save $doc 'formatting-rich.doc'

    # ---- formatting-styles.doc: custom paragraph and character styles ----
    $doc = $app.Documents.Add()
    $doc.Content.Text = "Heading text`rBody in custom style`rSecond custom paragraph`rNormal paragraph with a styled word inside."
    $my = $doc.Styles.Add('FixtureBody', 1)
    $my.BaseStyle = 'Normal'
    $my.Font.Name = 'Verdana'
    $my.Font.Size = 14
    $my.Font.Bold = -1
    $my.Font.Color = (Rgb 0 0 160)
    $my.ParagraphFormat.SpaceAfter = 10
    $my.ParagraphFormat.LeftIndent = 24
    $my.ParagraphFormat.Alignment = 1
    $my2 = $doc.Styles.Add('FixtureBodyTwo', 1)
    $my2.BaseStyle = 'FixtureBody'
    $my2.Font.Italic = -1
    $my2.Font.Bold = 0
    $my.NextParagraphStyle = 'FixtureBodyTwo'
    $ch = $doc.Styles.Add('FixtureEmphasis', 2)
    $ch.Font.Italic = -1
    $ch.Font.Color = (Rgb 0 128 0)
    $ch.Font.Underline = 1
    $h = $doc.Styles.Item(-2)
    $h.Font.Name = 'Georgia'
    $h.Font.Size = 20
    $h.Font.Color = (Rgb 128 0 0)
    $doc.Paragraphs.Item(1).Style = $doc.Styles.Item(-2)
    $doc.Paragraphs.Item(2).Style = 'FixtureBody'
    $doc.Paragraphs.Item(3).Style = 'FixtureBodyTwo'
    $q = $doc.Paragraphs.Item(4).Range
    $i = $q.Text.IndexOf('styled')
    $doc.Range($q.Start + $i, $q.Start + $i + 6).Style = 'FixtureEmphasis'
    Save $doc 'formatting-styles.doc'

    # ---- formatting-lists.doc: bullet, numbered and multilevel lists ----
    $doc = $app.Documents.Add()
    $doc.Content.Text = "Bullet one`rBullet two`rNumber one`rNumber two`rLevel one`rLevel two`rLevel three`rLevel one again`rPlain closing paragraph."
    $bul = $doc.ListTemplates.Add($false, 'FixtureBullets')
    $lv = $bul.ListLevels.Item(1)
    $lv.NumberFormat = [string][char]0xF0B7
    $lv.NumberStyle = 23
    $lv.TrailingCharacter = 0
    $lv.NumberPosition = 18
    $lv.TextPosition = 36
    $lv.TabPosition = 36
    $lv.Font.Name = 'Symbol'
    $doc.Range($doc.Paragraphs.Item(1).Range.Start, $doc.Paragraphs.Item(2).Range.End).ListFormat.ApplyListTemplate($bul, $false, 0)
    $num = $doc.ListTemplates.Add($false, 'FixtureNumbers')
    $lv = $num.ListLevels.Item(1)
    $lv.NumberFormat = '%1)'
    $lv.NumberStyle = 4
    $lv.StartAt = 3
    $lv.NumberPosition = 0
    $lv.TextPosition = 30
    $lv.TabPosition = 30
    $doc.Range($doc.Paragraphs.Item(3).Range.Start, $doc.Paragraphs.Item(4).Range.End).ListFormat.ApplyListTemplate($num, $false, 0)
    $multi = $doc.ListTemplates.Add($true, 'FixtureMulti')
    $spec = @(@('%1.', 0, 0, 36), @('%1.%2', 0, 36, 72), @('(%3)', 2, 72, 108))
    for ($n = 1; $n -le 3; $n++) {
        $lv = $multi.ListLevels.Item($n)
        $lv.NumberFormat = $spec[$n - 1][0]
        $lv.NumberStyle = $spec[$n - 1][1]
        $lv.NumberPosition = $spec[$n - 1][2]
        $lv.TextPosition = $spec[$n - 1][3]
        $lv.TabPosition = $spec[$n - 1][3]
        $lv.StartAt = 1
    }
    $doc.Range($doc.Paragraphs.Item(5).Range.Start, $doc.Paragraphs.Item(8).Range.End).ListFormat.ApplyListTemplate($multi, $false, 0)
    $doc.Paragraphs.Item(6).Range.ListFormat.ListLevelNumber = 2
    $doc.Paragraphs.Item(7).Range.ListFormat.ListLevelNumber = 3
    Save $doc 'formatting-lists.doc'

    # ---- formatting-unicode.doc: text edited after a first save, with a CJK character (Unicode piece) ----
    # Word 16 ignores Options.AllowFastSave through COM (a 64 KB document with a small edit is still
    # saved as one piece), so a genuine multi-piece file is produced in the tests with ole2's own
    # piece-append editor instead.
    $doc = $app.Documents.Add()
    $doc.Content.Text = "First fast paragraph.`rSecond fast paragraph.`rThird fast paragraph."
    $doc.SaveAs2((Join-Path $out 'formatting-unicode.doc'), 0, $false, '', $false)
    $doc.Range($doc.Paragraphs.Item(2).Range.Start, $doc.Paragraphs.Item(2).Range.Start).InsertBefore("Inserted ")
    $doc.Range($doc.Paragraphs.Item(1).Range.Start, $doc.Paragraphs.Item(1).Range.Start + 5).Font.Bold = -1
    $doc.Content.InsertAfter("Appended " + [char]0x00E9 + [char]0x4E2D + " fast text.`r")
    $doc.Save()
    $doc.Close($false)
    Write-Output 'done'
}
finally {
    $app.Quit()
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
