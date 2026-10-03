param([string]$OutputDirectory = $PSScriptRoot)
# Owned synthetic numeric-record fixture. No source workbooks or personal files.
# pwsh -STA -NoProfile -File this-script.ps1
$ErrorActionPreference = 'Stop'
$out = [IO.Path]::GetFullPath($OutputDirectory)
[void][IO.Directory]::CreateDirectory($out)
$excel = $null
$book = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AutomationSecurity = 3
    $excel.EnableEvents = $false
    $book = $excel.Workbooks.Add()
    while ($book.Worksheets.Count -gt 1) { $book.Worksheets.Item($book.Worksheets.Count).Delete() }
    $sheet = $book.Worksheets.Item(1)
    $sheet.Name = 'Numbers'
    foreach ($column in 1..4) { $sheet.Cells.Item(1, $column).Value2 = $column * 10 }
    # Different XF entries inside the same contiguous numeric RK region.
    $sheet.Range('A1').NumberFormat = '0.00'
    $sheet.Range('B1').Font.Bold = $true
    $sheet.Range('C1').Interior.Color = 65535
    $sheet.Range('D1').Font.Italic = $true
    $sheet.Range('A2').Value2 = 'Alias anchor'
    $sheet.Range('B2').Value2 = 'Alias anchor'
    $sheet.Range('C2').Value2 = 'Untouched anchor'
    foreach ($column in 1..4) { $sheet.Cells.Item(3, $column).Value2 = $column * 100 }
    $sheet.Range('A4').Formula = '=SUM(A1:D1)'
    $sheet.Range('B4').Value2 = -1234567.891
    $sheet.Range('A2').AddComment('Synthetic comment unchanged') | Out-Null
    $sheet.Range('D5:E5').Merge()
    $sheet.Range('D5').Value2 = 'Merged anchor'
    $sheet.Columns.Item(1).ColumnWidth = 22
    $sheet.Columns.Item(2).ColumnWidth = 17
    $props = $book.BuiltinDocumentProperties
    foreach ($entry in @(@('Author', 'Fixture Author'), @('Title', 'Synthetic MULRK conversion fixture'))) {
        $prop = [System.__ComObject].InvokeMember('Item', 'GetProperty', $null, $props, @($entry[0]))
        [void][System.__ComObject].InvokeMember('Value', 'SetProperty', $null, $prop, @($entry[1]))
    }
    $book.RemovePersonalInformation = $true
    $book.SaveAs((Join-Path $out 'workbook-mulrk.xls'), 56)
    Write-Output (Join-Path $out 'workbook-mulrk.xls')
}
finally {
    try {
        if ($null -ne $book) { $book.Close($false); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($book) }
    }
    finally {
        if ($null -ne $excel) { $excel.Quit(); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) }
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}
