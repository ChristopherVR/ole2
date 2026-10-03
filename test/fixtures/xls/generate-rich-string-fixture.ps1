# Synthetic public test workbook; no personal files read.
# Run with Windows PowerShell STA: powershell.exe -NoProfile -Sta -File this.ps1
# Desktop Excel saves native BIFF8 (FileFormat 56). Original content is Apache-2.0.
$ErrorActionPreference = 'Stop'
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.AutomationSecurity = 3
$wb = $null
try {
 $wb = $excel.Workbooks.Add()
 while ($wb.Worksheets.Count -gt 2) { $wb.Worksheets.Item($wb.Worksheets.Count).Delete() }
 while ($wb.Worksheets.Count -lt 2) { [void]$wb.Worksheets.Add([Type]::Missing,$wb.Worksheets.Item($wb.Worksheets.Count)) }
 $s = $wb.Worksheets.Item(1)
 $s.Name = 'Aliases'
 $s.Range('A1').Value2 = 'Shared plain text'
 $s.Range('B1').Value2 = 'Shared plain text'
 $s.Range('A2').Value2 = ('Rich preserved ' + [char]0x03a9)
 $s.Range('A2').Characters(1,4).Font.Bold = $true
 $s.Range('A2').Characters(6,9).Font.Italic = $true
 $s.Range('B2').Value2 = ('Rich preserved ' + [char]0x03a9)
 $s.Range('A3').Value2 = ('0123456789' * 1000)
 $s.Range('A4').Value2 = (([string][char]0x65e5 + [char]0x672c + [char]0x8a9e + [char]0x03a9) * 3000)
 $s.Range('A5').Value2 = 'Replace me'
 $s.Range('B1').AddComment('An untouched native comment') | Out-Null
 $s.Range('D1:E1').Merge()
 $s.Range('D1').Value2 = 'Merged'
 $s.Rows.Item(2).RowHeight = 22
 $s.Columns.Item(1).ColumnWidth = 30
 $later = $wb.Worksheets.Item(2)
 $later.Name = 'Later'
 $later.Range('A1').Value2 = 'Later edit'
 $later.Range('A2').Value2 = 42
 $later.Range('A3').Formula = "=Aliases!B1&`" preserved`""
 $props = $wb.BuiltinDocumentProperties
 $author = [System.__ComObject].InvokeMember('Item','GetProperty',$null,$props,@('Author'))
 [void][System.__ComObject].InvokeMember('Value','SetProperty',$null,$author,@('Fixture Author'))
 $wb.RemovePersonalInformation = $true
 $path = Join-Path $PSScriptRoot 'workbook-rich-strings.xls'
 $wb.SaveAs($path,56)
 $wb.Close($false)
 $wb = $null
 Write-Output "Generated $path"
} finally {
 if ($null -ne $wb) { try { $wb.Close($false) } catch {} }
 for ($i=0;$i -lt 20;$i++) { try { $excel.Quit(); break } catch { Start-Sleep -Milliseconds 100 } }
 [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)
}
