param([string]$OutputDirectory = $PSScriptRoot)
# Repository-owned synthetic BOOLERR and formula-neighbor fixture.
# pwsh -STA; no source documents, links, objects, macros or external formulas.
$ErrorActionPreference = 'Stop'
$out = [IO.Path]::GetFullPath($OutputDirectory)
[void][IO.Directory]::CreateDirectory($out)
$app = $null
$book = $null
try {
    $app = New-Object -ComObject Excel.Application
    $app.Visible = $false
    $app.DisplayAlerts = $false
    $app.AutomationSecurity = 3
    $app.EnableEvents = $false
    $book = $app.Workbooks.Add()
    while ($book.Worksheets.Count -gt 1) { $book.Worksheets.Item($book.Worksheets.Count).Delete() }
    $sheet = $book.Worksheets.Item(1)
    $sheet.Name = 'Types'
    $sheet.Range('A1').Value2 = $true
    $sheet.Range('B1').Value2 = $false
    # Assign Excel's native evaluated error values, then remove the formulas.
    # These constant formulas have no external references or executable content.
    $sheet.Range('C1').Formula = '=1/0'
    $sheet.Range('D1').Formula = '=NA()'
    $sheet.Range('C1').Copy()
    $sheet.Range('C1').PasteSpecial(-4163) # xlPasteValues, preserves VT_ERROR
    $sheet.Range('D1').Copy()
    $sheet.Range('D1').PasteSpecial(-4163)
    $app.CutCopyMode = $false
    $sheet.Range('A2').Value2 = 10
    $sheet.Range('B2').Value2 = 'Untouched literal'
    $sheet.Range('C2').Formula = '=SUM(A2,5)'
    $sheet.Range('D2').Formula = '=IF(A1,11,22)'
    $sheet.Columns.ColumnWidth = 22
    $props = $book.BuiltinDocumentProperties
    foreach ($entry in @(@('Author', 'Fixture Author'), @('Title', 'Synthetic boolean and error fixture'))) {
        $prop = [System.__ComObject].InvokeMember('Item', 'GetProperty', $null, $props, @($entry[0]))
        [void][System.__ComObject].InvokeMember('Value', 'SetProperty', $null, $prop, @($entry[1]))
    }
    $book.RemovePersonalInformation = $true
    $book.SaveAs((Join-Path $out 'workbook-cell-types.xls'), 56)
    Write-Output (Join-Path $out 'workbook-cell-types.xls')
}
finally {
    try { if ($null -ne $book) { $book.Close($false); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($book) } }
    finally {
        if ($null -ne $app) { $app.Quit(); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app) }
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}
