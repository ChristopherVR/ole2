param([string]$OutputDirectory=$PSScriptRoot)
# Original neutral content; no source workbooks or personal documents.
# pwsh -STA -NoProfile -File test/fixtures/xls/generate-blank-fixture.ps1
# https://learn.microsoft.com/en-us/office/vba/api/excel.xlfileformat
$ErrorActionPreference='Stop'
$out=[IO.Path]::GetFullPath($OutputDirectory)
[void][IO.Directory]::CreateDirectory($out)
$app=$null;$book=$null
try {
    $app=New-Object -ComObject Excel.Application
    $app.Visible=$false;$app.DisplayAlerts=$false
    $app.AutomationSecurity=3;$app.EnableEvents=$false
    $book=$app.Workbooks.Add()
    while($book.Worksheets.Count -gt 2){$book.Worksheets.Item($book.Worksheets.Count).Delete()}
    while($book.Worksheets.Count -lt 2){[void]$book.Worksheets.Add([Type]::Missing,$book.Worksheets.Item($book.Worksheets.Count))}
    $sheet=$book.Worksheets.Item(1);$sheet.Name='Blanks'
    # A row containing only formatted empty cells, distinct XFs first/middle/last.
    $sheet.Range('A1').NumberFormat='0.00'
    $sheet.Range('B1').Font.Bold=$true
    $sheet.Range('C1').Font.Italic=$true
    $sheet.Range('D1').Font.Size=16
    # Both isolated and adjacent formatted blanks between nonblank records.
    $sheet.Range('A2').Value2='Mixed left anchor'
    $sheet.Range('B2').Font.Bold=$true
    $sheet.Range('C2').Value2=42
    $sheet.Range('D2').NumberFormat='0.000'
    $sheet.Range('E2').Font.Italic=$true
    $sheet.Range('F2').Font.Size=18
    $sheet.Range('G2').Value2=73
    $sheet.Range('H2').Value2='Mixed right anchor'
    foreach($column in 1..4){$sheet.Cells.Item(3,$column).Value2=$column*100}
    $sheet.Range('A4').Formula='=SUM(A3:D3)'
    $sheet.Range('B4').Value2='Shared anchor'
    $sheet.Range('C4').Value2='Shared anchor'
    $sheet.Range('A5').Value2='Comment anchor'
    $sheet.Range('A5').AddComment('Owned synthetic comment unchanged')|Out-Null
    $sheet.Range('C5:D5').Merge();$sheet.Range('C5').Value2='Merged anchor'
    $sheet.Columns.Item(1).ColumnWidth=24;$sheet.Columns.Item(2).ColumnWidth=17
    $sheet.Rows.Item(1).RowHeight=23
    $later=$book.Worksheets.Item(2);$later.Name='Later'
    $later.Range('A1').Value2='Later text anchor'
    $later.Range('A2').Value2=12
    $later.Range('B2').Font.Bold=$true
    $later.Range('A3').Formula='=Blanks!A3+Blanks!B3'
    $props=$book.BuiltinDocumentProperties
    foreach($entry in @(@('Author','Fixture Author'),@('Title','Owned stored blank cell fixture'),@('Company','Fixture Organization'))){
        $property=[System.__ComObject].InvokeMember('Item','GetProperty',$null,$props,@($entry[0]))
        [void][System.__ComObject].InvokeMember('Value','SetProperty',$null,$property,@($entry[1]))
    }
    $book.RemovePersonalInformation=$true
    $book.SaveAs((Join-Path $out 'workbook-blanks.xls'),56)
    Write-Output (Join-Path $out 'workbook-blanks.xls')
} finally {
    try{if($null -ne $book){$book.Close($false);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($book)}}finally{
        if($null -ne $app){$app.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
        [GC]::Collect();[GC]::WaitForPendingFinalizers()
    }
}
