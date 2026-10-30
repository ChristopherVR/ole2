param([string]$OutputDirectory=$PSScriptRoot)
# Builds neutral synthetic content only (no personal documents) through Word COM and saves Word 97-2003 files.
# Values set here are asserted by test/doc-structure.test.ts. No global Word options are changed.
$ErrorActionPreference='Stop'
$out=[IO.Path]::GetFullPath($OutputDirectory);[void][IO.Directory]::CreateDirectory($out)
function Set-Prop($doc,$name,$value){
    $props=$doc.BuiltInDocumentProperties
    $bt=[System.__ComObject]
    $p=$bt.InvokeMember('Item','GetProperty',$null,$props,@($name))
    $bt.InvokeMember('Value','SetProperty',$null,$p,@($value))
}
function Save-Doc($app,$doc,$name){
    Set-Prop $doc 'Author' 'Fixture Author'
    $pth=[string](Join-Path $out $name); $doc.SaveAs2($pth,0)
    Write-Output (Join-Path $out $name)
}
function New-Doc($app){ $d=$app.Documents.Add(); $d }
function Rgb($r,$g,$b){ $r + 256*$g + 65536*$b }
$app=New-Object -ComObject Word.Application
$app.Visible=$false;$app.DisplayAlerts=0;$app.AutomationSecurity=3
try {
    # 1. simple table
    $doc=New-Doc $app
    $doc.Content.Text='Before table'
    $doc.Content.InsertParagraphAfter()
    $r=$doc.Paragraphs.Item($doc.Paragraphs.Count).Range
    $t=$doc.Tables.Add($r,3,2)
    for($i=1;$i -le 3;$i++){for($j=1;$j -le 2;$j++){$t.Cell($i,$j).Range.Text="R${i}C${j}"}}
    $t.Borders.Enable=$true
    $t.Columns.Item(1).Width=100; $t.Columns.Item(2).Width=150
    $t.Rows.Item(1).HeadingFormat=$true
    $t.Rows.Item(2).HeightRule=2; $t.Rows.Item(2).Height=30
    $t.Rows.Item(3).AllowBreakAcrossPages=$false
    $t.Rows.Alignment=1
    $t.Cell(2,1).Shading.BackgroundPatternColor=(Rgb 255 0 0)
    $t.Cell(3,2).VerticalAlignment=1
    $doc.Content.InsertParagraphAfter()
    $doc.Paragraphs.Item($doc.Paragraphs.Count).Range.Text='After table'
    [void](Save-Doc $app $doc 'structure-table-simple.doc'); $doc.Close($false)

    # 2. merged cells
    $doc=New-Doc $app
    $doc.Content.Text='Merged'
    $doc.Content.InsertParagraphAfter()
    $r=$doc.Paragraphs.Item($doc.Paragraphs.Count).Range
    $t=$doc.Tables.Add($r,3,3)
    for($i=1;$i -le 3;$i++){for($j=1;$j -le 3;$j++){$t.Cell($i,$j).Range.Text="R${i}C${j}"}}
    $t.Cell(1,1).Merge($t.Cell(1,2))
    $t.Cell(2,3).Merge($t.Cell(3,3))
    [void](Save-Doc $app $doc 'structure-table-merged.doc'); $doc.Close($false)

    # 3. nested table
    $doc=New-Doc $app
    $doc.Content.Text='Nested'
    $doc.Content.InsertParagraphAfter()
    $r=$doc.Paragraphs.Item($doc.Paragraphs.Count).Range
    $t=$doc.Tables.Add($r,2,2)
    $t.Cell(1,1).Range.Text='Outer11'; $t.Cell(2,1).Range.Text='Outer21'; $t.Cell(2,2).Range.Text='Outer22'
    $t.Cell(1,2).Range.Text='Lead'
    $cr=$t.Cell(1,2).Range; $cr.Collapse(0); $cr.MoveEnd(1,-1) | Out-Null
    $cr=$t.Cell(1,2).Range; $cr.InsertParagraphAfter()
    $cr=$t.Cell(1,2).Range; $cr.MoveEnd(1,-1) | Out-Null; $cr.Collapse(0)
    $n=$doc.Tables.Add($cr,2,2)
    $n.Cell(1,1).Range.Text='In11';$n.Cell(1,2).Range.Text='In12';$n.Cell(2,1).Range.Text='In21';$n.Cell(2,2).Range.Text='In22'
    [void](Save-Doc $app $doc 'structure-table-nested.doc'); $doc.Close($false)

    # 4. sections
    $doc=New-Doc $app
    $doc.Content.Text='Section one'
    $e=$doc.Content; $e.Collapse(0); $e.InsertBreak(2)
    $e=$doc.Content; $e.Collapse(0); $e.InsertAfter('Section two'); 
    $e=$doc.Content; $e.Collapse(0); $e.InsertBreak(3)
    $e=$doc.Content; $e.Collapse(0); $e.InsertAfter('Section three')
    $e=$doc.Content; $e.Collapse(0); $e.InsertBreak(3)
    $e=$doc.Content; $e.Collapse(0); $e.InsertAfter('Section four')
    $s1=$doc.Sections.Item(1).PageSetup
    $s1.LeftMargin=72; $s1.RightMargin=72; $s1.TopMargin=72; $s1.BottomMargin=72; $s1.HeaderDistance=36; $s1.FooterDistance=36
    $s2=$doc.Sections.Item(2).PageSetup
    $s2.Orientation=1; $s2.LeftMargin=54; $s2.RightMargin=90; $s2.TopMargin=60; $s2.BottomMargin=66; $s2.HeaderDistance=30; $s2.FooterDistance=40; $s2.Gutter=18
    $s2.TextColumns.SetCount(2); $s2.TextColumns.Spacing=30; $s2.LineNumbering.Active=0
    $s3=$doc.Sections.Item(3).PageSetup
    $s3.TextColumns.SetCount(1)
    $pn=$doc.Sections.Item(3).Footers.Item(1).PageNumbers
    $pn.RestartNumberingAtSection=$true; $pn.StartingNumber=5; $pn.NumberStyle=2
    $doc.Sections.Item(2).PageSetup.SectionStart=4; $doc.Sections.Item(3).PageSetup.SectionStart=2; $doc.Sections.Item(4).PageSetup.SectionStart=0
    [void](Save-Doc $app $doc 'structure-sections.doc'); $doc.Close($false)

    # 5. headers and footers
    $doc=New-Doc $app
    $doc.Content.Text='Section A'
    $e=$doc.Content; $e.Collapse(0); $e.InsertBreak(2)
    $e=$doc.Content; $e.Collapse(0); $e.InsertAfter('Section B')
    $ps=$doc.Sections.Item(1).PageSetup
    $ps.OddAndEvenPagesHeaderFooter=-1; $ps.DifferentFirstPageHeaderFooter=-1
    $s=$doc.Sections.Item(1)
    $s.Headers.Item(1).Range.Text='Odd header 1';  $s.Headers.Item(2).Range.Text='First header 1'; $s.Headers.Item(3).Range.Text='Even header 1'
    $s.Footers.Item(1).Range.Text='Odd footer 1';  $s.Footers.Item(2).Range.Text='First footer 1'; $s.Footers.Item(3).Range.Text='Even footer 1'
    $s2=$doc.Sections.Item(2)
    $s2.PageSetup.DifferentFirstPageHeaderFooter=0
    $s2.Headers.Item(1).LinkToPrevious=$false
    $s2.Headers.Item(1).Range.Text='Odd header 2'
    [void](Save-Doc $app $doc 'structure-headers.doc'); $doc.Close($false)

    # 6. footnotes, endnotes, comments
    $doc=New-Doc $app
    $doc.Content.Text='Alpha beta gamma delta epsilon zeta eta theta.'
    $rg=$doc.Range(0,5); [void]$doc.Footnotes.Add($doc.Range(5,5),[Type]::Missing,'First footnote text')
    [void]$doc.Footnotes.Add($doc.Range(16,16),[Type]::Missing,'Second footnote text')
    [void]$doc.Endnotes.Add($doc.Range(22,22),[Type]::Missing,'First endnote text')
    [void]$doc.Endnotes.Add($doc.Range(30,30),[Type]::Missing,'Second endnote text')
    $c1=$doc.Comments.Add($doc.Range(0,5),'First comment text'); $c1.Author='Alice Author'; $c1.Initial='AA'
    $c2=$doc.Comments.Add($doc.Range(6,10),'Second comment text'); $c2.Author='Bob Builder'; $c2.Initial='BB'
    [void](Save-Doc $app $doc 'structure-notes.doc'); $doc.Close($false)

    # 7. text boxes
    $doc=New-Doc $app
    $doc.Content.Text='Body text with boxes.'
    $a=$doc.Range(0,0)
    $b1=$doc.Shapes.AddTextbox(1,100,100,200,50,$a); $b1.TextFrame.TextRange.Text='Text box one'
    $a2=$doc.Range(10,10)
    $b2=$doc.Shapes.AddTextbox(1,150,200,180,40,$a2); $b2.TextFrame.TextRange.Text='Text box two'
    [void](Save-Doc $app $doc 'structure-textbox.doc'); $doc.Close($false)
} finally {
    try{foreach($d in @($app.Documents)){$d.Close($false)}}catch{}
    $app.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)
    [GC]::Collect();[GC]::WaitForPendingFinalizers()
}
