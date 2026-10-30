param([string]$OutputDirectory = $PSScriptRoot)
# Owned synthetic fixture for pictures, shapes, fields, bookmarks and an embedded object.
# Requires desktop Word and System.Drawing: pwsh -STA -NoProfile -File this-script.ps1
# Writes objects-input-{a.png,b.jpg} (deterministic generated images) and objects.doc (FileFormat 0).
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Drawing
$out = [IO.Path]::GetFullPath($OutputDirectory)
[void][IO.Directory]::CreateDirectory($out)
function New-Image([string]$path, [int]$w, [int]$h, $format, $c1, $c2) {
    $bmp = New-Object System.Drawing.Bitmap $w, $h
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.Clear($c1)
    $brush = New-Object System.Drawing.SolidBrush $c2
    $g.FillRectangle($brush, [int]($w / 2), 0, [int]($w / 2), [int]($h / 2))
    $g.Dispose(); $brush.Dispose()
    $bmp.Save($path, $format); $bmp.Dispose()
}
$png = Join-Path $out 'objects-input-a.png'
$jpg = Join-Path $out 'objects-input-b.jpg'
New-Image $png 64 48 ([System.Drawing.Imaging.ImageFormat]::Png) ([System.Drawing.Color]::FromArgb(200, 30, 30)) ([System.Drawing.Color]::FromArgb(30, 60, 200))
New-Image $jpg 80 40 ([System.Drawing.Imaging.ImageFormat]::Jpeg) ([System.Drawing.Color]::FromArgb(20, 160, 60)) ([System.Drawing.Color]::FromArgb(240, 220, 20))
$app = $null; $doc = $null
try {
    $app = New-Object -ComObject Word.Application
    $app.Visible = $false; $app.DisplayAlerts = 0; $app.AutomationSecurity = 3
    $doc = $app.Documents.Add()
    $doc.Content.Text = "Inline PNG: X and inline JPEG: Y end.`rLink: HLINK anchor ANCH.`rBookmarked text here.`rFields: F1 F2 F3 F4 F5 F6.`rFloating anchor paragraph.`rEmbedded: OBJ done.`rTOC:`r"
    function Place([int]$p, [string]$token) {
        $r = $doc.Paragraphs.Item($p).Range
        $f = $r.Find
        [void]$f.Execute($token)
        return $r
    }
    # Inline pictures
    $r = Place 1 'X'; $r.Text = ''; $is1 = $doc.InlineShapes.AddPicture($png, $false, $true, $r)
    $is1.Width = 96; $is1.Height = 72; $is1.AlternativeText = 'inline png'
    $r = Place 1 'Y'; $r.Text = ''; $is2 = $doc.InlineShapes.AddPicture($jpg, $false, $true, $r)
    # Bookmarks
    $r = $doc.Paragraphs.Item(3).Range; $r.End = $r.End - 1
    [void]$doc.Bookmarks.Add('BM_Target', $r)
    $r = Place 2 'ANCH'; [void]$doc.Bookmarks.Add('BM_Point', $r)
    # Hyperlinks
    $r = Place 2 'HLINK'; [void]$doc.Hyperlinks.Add($r, 'https://example.com/path?x=1', '', 'Example tip', 'ext link')
    $r = Place 2 'ANCH'; [void]$doc.Hyperlinks.Add($r, '', 'BM_Target', 'Local tip', 'to bookmark')
    # Fields
    $fields = @(@('F1', 33, ''), @('F2', 26, ''), @('F3', 31, '\@ "yyyy-MM-dd"'), @('F4', -1, 'SEQ Figure \* ARABIC'), @('F5', -1, 'MERGEFIELD CustomerName \* MERGEFORMAT'), @('F6', -1, 'INCLUDEPICTURE "C:\nonexistent\pic.png"'))
    foreach ($e in $fields) {
        $r = Place 4 $e[0]; $r.Text = ''
        if ($e[1] -eq -1) { [void]$doc.Fields.Add($r, -1, $e[2], $false) } else { [void]$doc.Fields.Add($r, $e[1], $e[2], $false) }
    }
    $r = $doc.Paragraphs.Item(3).Range; $r.Collapse(0); $r.MoveEnd(1, -1) | Out-Null
    $r = $doc.Paragraphs.Item(3).Range; $r.Collapse(1)
    [void]$doc.Fields.Add($r, -1, 'REF BM_Target \h', $false)
    $r = $doc.Paragraphs.Item(3).Range; $r.Collapse(1)
    [void]$doc.Fields.Add($r, -1, 'PAGEREF BM_Target \h', $false)
    # Floating shapes
    $anchor = $doc.Paragraphs.Item(5).Range
    $s1 = $doc.Shapes.AddPicture($png, $false, $true, 300, 120, 100, 75, $anchor)
    $s1.WrapFormat.Type = 0; $s1.Name = 'FloatPng'
    $s2 = $doc.Shapes.AddPicture($jpg, $false, $true, 100, 200, 80, 40, $anchor)
    $s2.WrapFormat.Type = 3; $s2.Name = 'FloatJpg'
    $s3 = $doc.Shapes.AddShape(1, 50, 300, 120, 60, $anchor)
    $s3.Fill.ForeColor.RGB = 0x0000FF; $s3.Line.ForeColor.RGB = 0x00FF00; $s3.Line.Weight = 3; $s3.Name = 'Rect'
    $s4 = $doc.Shapes.AddTextbox(1, 250, 300, 140, 50, $anchor)
    $s4.TextFrame.TextRange.Text = 'Textbox content'; $s4.Name = 'Tbox'
    # Embedded object
    $r = Place 6 'OBJ'; $r.Text = ''
    # Excel is installed alongside Word on the generating machine; a new empty worksheet object.
    $o = $doc.InlineShapes.AddOLEObject('Excel.Sheet.8', [Type]::Missing, $false, $false, [Type]::Missing, [Type]::Missing, [Type]::Missing, $r)
    # TOC
    $p = $doc.Paragraphs.Item(1).Range; $p.Style = $doc.Styles.Item(-2)
    $r = $doc.Paragraphs.Item(7).Range; $r.Collapse(1)
    [void]$doc.TablesOfContents.Add($r, $true, 1, 1)
    $doc.RemovePersonalInformation = $true; $doc.EmbedTrueTypeFonts = $false
    $doc.SaveAs2((Join-Path $out 'objects.doc'), 0, $false, '', $false)
    Write-Output (Join-Path $out 'objects.doc')
}
finally {
    try { if ($null -ne $doc) { $doc.Close($false); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($doc) } }
    finally {
        if ($null -ne $app) { $app.Quit(); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app) }
        [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    }
}
