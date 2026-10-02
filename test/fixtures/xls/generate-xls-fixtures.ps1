# Regenerates the BIFF8 .xls fixtures used by test/legacy-excel-workbook*.test.ts.
# Requires desktop Excel (COM automation). Run from any directory:
#   pwsh -File test/fixtures/xls/generate-xls-fixtures.ps1
# Every file is saved with FileFormat 56 (xlExcel8, Excel 97-2003 workbook).
$ErrorActionPreference = 'Stop'
$out = $PSScriptRoot
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.UserName = 'Fixture Author'

function Set-DocProperty($wb, $name, $value) {
	$props = $wb.BuiltinDocumentProperties
	$prop = [System.__ComObject].InvokeMember('Item', 'GetProperty', $null, $props, @($name))
	[void][System.__ComObject].InvokeMember('Value', 'SetProperty', $null, $prop, @($value))
}

function Save-Xls($wb, $name, $password) {
	$path = Join-Path $out $name
	# Keep personal account details out of committed fixtures.
	Set-DocProperty $wb 'Author' 'Fixture Author'
	Set-DocProperty $wb 'Title' ('Fixture ' + $name)
	$wb.RemovePersonalInformation = $true
	if (Test-Path $path) { Remove-Item $path -Force }
	if ($password) {
		$wb.SaveAs($path, 56, $password)
	}
	else {
		$wb.SaveAs($path, 56)
	}
	$wb.Close($false)
}

function New-Book($count) {
	$wb = $excel.Workbooks.Add()
	while ($wb.Worksheets.Count -lt $count) { [void]$wb.Worksheets.Add([Type]::Missing, $wb.Worksheets.Item($wb.Worksheets.Count)) }
	while ($wb.Worksheets.Count -gt $count) { $wb.Worksheets.Item($wb.Worksheets.Count).Delete() }
	return $wb
}

try {
	# 1. workbook-features.xls: values, formulas, hidden sheets, long strings.
	$wb = New-Book 4
	$data = $wb.Worksheets.Item(1)
	$data.Name = 'Data'
	# Name every sheet before writing formulas so cross-sheet references resolve.
	$wb.Worksheets.Item(2).Name = 'My Formulas'
	$wb.Worksheets.Item(3).Name = 'Hidden'
	$wb.Worksheets.Item(4).Name = 'Secret'
	$data.Range('A1').Value2 = 'Name'
	$data.Range('B1').Value2 = 'Price'
	$data.Range('C1').Value2 = 'Qty'
	$data.Range('A2').Value2 = 'Apple'
	$data.Range('B2').Value2 = 1.25
	$data.Range('C2').Value2 = 10
	$data.Range('A3').Value2 = 'Banana'
	$data.Range('B3').Value2 = 0.5
	$data.Range('C3').Value2 = 24
	$data.Range('A4').Value2 = 'Cherry'
	$data.Range('B4').Value2 = 3.75
	$data.Range('C4').Value2 = 7
	$data.Range('A5').Value2 = ([string][char]0x65E5 + [char]0x672C + [char]0x8A9E + ' ' + [char]0x00E9 + 't' + [char]0x00E9 + ' ' + [char]0x0416)
	$data.Range('B5').Value2 = -1234567.891
	$data.Range('C5').Value2 = 123456789
	$data.Range('D1').Value2 = 'Flag'
	$data.Range('D2').Value2 = $true
	$data.Range('D3').Value2 = $false
	$data.Range('E2').Formula = '=1/0'
	$data.Range('E3').Formula = '=NA()'
	$data.Range('F2').Value2 = 45292
	$data.Range('F2').NumberFormat = 'yyyy-mm-dd'
	$data.Range('F3').Value2 = 0.75
	$data.Range('F3').NumberFormat = 'h:mm AM/PM'
	$data.Range('G2').Value2 = 0.125
	$data.Range('G2').NumberFormat = '0.00%'
	# A narrow string longer than one SST record (8224 bytes) forces CONTINUE.
	$data.Range('A8').Value2 = ('0123456789' * 1000)
	# A wide (UTF-16) string that also spans CONTINUE records.
	$data.Range('A9').Value2 = (([string][char]0x00C5 + [char]0x03A9 + [char]0x4E2D + 'x') * 3000)
	# Many short strings so the SST itself spans records.
	for ($i = 1; $i -le 600; $i++) { $data.Cells.Item(10 + $i, 1).Value2 = "Item number $i" }

	$formulas = $wb.Worksheets.Item(2)
	$formulas.Name = 'My Formulas'
	$formulas.Range('A1').Formula = '=SUM(Data!B2:B4)'
	$formulas.Range('A2').Formula = '=IF(Data!B2>1,"big","small")'
	$formulas.Range('A3').Formula = '=VLOOKUP("Banana",Data!A2:C4,3,FALSE)'
	$formulas.Range('A4').Formula = '=Data!B2*Data!C2+1'
	$formulas.Range('A5').Formula = '=CONCATENATE(Data!A2," & ",Data!A3)'
	$formulas.Range('A6').Formula = '=ROUND(AVERAGE(Data!B2:B4),2)'
	$formulas.Range('A7').Formula = '=A1-A4/2'
	$formulas.Range('A8').Formula = '=COUNTIF(Data!C2:C4,">8")'
	$formulas.Range('A9').Formula = '=$A$1+A$2&"x"'
	$formulas.Range('A10').Formula = '=MAX(1,2,3)-MIN(A1:A4)'
	$formulas.Range('A11').Formula = '=Data!D2'
	$formulas.Range('A12').Formula = '=ISERROR(Data!E2)'
	$formulas.Range('A13').Formula = '=-A1^2%'
	$formulas.Range('A14').Formula = '=TODAY()-TODAY()'
	$formulas.Range('A15').Formula = '=IF(A1>0,SUM(A1:A2),"")'
	$formulas.Range('A16').Formula = '=Data!E2'
	$formulas.Range('A17').Formula = '="a"&"b"'
	$formulas.Range('A18').Formula = '=CHOOSE(2,"x","y","z")'
	$formulas.Range('A19').Formula = '=SUM(Data!B:B)'
	$formulas.Range('A20').Formula = '=Hidden!A1'
	$formulas.Range('A21').Formula = '=AND(TRUE,NOT(FALSE))'
	$formulas.Range('A22').Formula = '=SUMPRODUCT(Data!B2:B4,Data!C2:C4)'
	$formulas.Range('A23').Formula = '=IFERROR(1/0,"err")'
	$formulas.Range('A24').Formula = '=SUMIFS(Data!B2:B4,Data!C2:C4,">8")'
	$formulas.Range('A25').Formula = '=EOMONTH(Data!F2,1)'
	$formulas.Range('A26').Formula = '=A1 + A4'
	$formulas.Range('A27').Formula = '=SUM({1,2;3,4})'
	$formulas.Range('A28').Formula = '=SUM(Data:Hidden!Z1:Z2)'
	$formulas.Range('A29').Formula = '=Secret!A1&" "&''My Formulas''!A2'
	# Shared formula: the same relative formula filled down a column.
	$formulas.Range('B1:B6').Formula = '=Data!B2*2'
	# Array formula.
	$formulas.Range('C1:C3').FormulaArray = '=Data!B2:B4*10'
	[void]$wb.Names.Add('TaxRate', '=Data!$B$2')
	$data.PageSetup.PrintArea = '$A$1:$C$5'
	$formulas.Range('D1').Formula = '=TaxRate*100'

	$hidden = $wb.Worksheets.Item(3)
	$hidden.Name = 'Hidden'
	$hidden.Range('A1').Value2 = 42
	$hidden.Visible = 0
	$veryHidden = $wb.Worksheets.Item(4)
	$veryHidden.Name = 'Secret'
	$veryHidden.Range('A1').Value2 = 'shh'
	$veryHidden.Visible = 2
	$data.Activate()
	Save-Xls $wb 'workbook-features.xls' $null

	# 2. workbook-styles.xls: formatting, merges, sizes, panes, comments, hyperlinks.
	$wb = New-Book 2
	$s = $wb.Worksheets.Item(1)
	$s.Name = 'Styled'
	$s.Range('A1').Value2 = 'Bold red'
	$s.Range('A1').Font.Bold = $true
	$s.Range('A1').Font.Color = 255
	$s.Range('A2').Value2 = 'Italic 14pt Arial'
	$s.Range('A2').Font.Italic = $true
	$s.Range('A2').Font.Size = 14
	$s.Range('A2').Font.Name = 'Arial'
	$s.Range('A3').Value2 = 'Yellow fill'
	$s.Range('A3').Interior.Color = 65535
	$s.Range('A4').Value2 = 'Underline strike'
	$s.Range('A4').Font.Underline = 2
	$s.Range('A4').Font.Strikethrough = $true
	$s.Range('B2').Value2 = 'Borders'
	$s.Range('B2').Borders.Item(7).LineStyle = 1
	$s.Range('B2').Borders.Item(7).Weight = 2
	$s.Range('B2').Borders.Item(8).LineStyle = 1
	$s.Range('B2').Borders.Item(8).Weight = -4138
	$s.Range('B2').Borders.Item(9).LineStyle = -4119
	$s.Range('B2').Borders.Item(9).Color = 16711680
	$s.Range('B2').Borders.Item(10).LineStyle = -4115
	$s.Range('B3').Value2 = 1234.5
	$s.Range('B3').NumberFormat = '#,##0.00'
	$s.Range('B4').Value2 = 1234.5
	$s.Range('B4').NumberFormat = '"USD" #,##0.000;[Red]-#,##0.000'
	$s.Range('B5').Value2 = 'Centered wrap text that is long'
	$s.Range('B5').HorizontalAlignment = -4108
	$s.Range('B5').VerticalAlignment = -4160
	$s.Range('B5').WrapText = $true
	$s.Range('B6').Value2 = 'Pattern'
	$s.Range('B6').Interior.Pattern = 9
	$s.Range('B6').Interior.PatternColor = 255
	$s.Range('B6').Interior.Color = 16777164
	$s.Range('C1').Value2 = 'Rotated'
	$s.Range('C1').Orientation = 45
	$s.Range('C2').Value2 = 'Indented'
	$s.Range('C2').IndentLevel = 2
	$s.Range('D1').Value2 = 'Merged block'
	$s.Range('D1:F2').Merge()
	$s.Range('D4:E4').Merge()
	$s.Columns.Item(1).ColumnWidth = 25
	$s.Columns.Item(3).ColumnWidth = 4.5
	$s.Columns.Item(7).Hidden = $true
	$s.Rows.Item(2).RowHeight = 30
	$s.Rows.Item(9).Hidden = $true
	$s.Range('A6').Value2 = 'Has comment'
	[void]$s.Range('A6').AddComment('Reviewed by QA')
	$s.Range('A7').Value2 = 'Link'
	[void]$s.Hyperlinks.Add($s.Range('A7'), 'https://example.com/path?q=1', '', 'Example tip', 'Example site')
	$s.Range('A8').Value2 = 'Jump'
	[void]$s.Hyperlinks.Add($s.Range('A8'), '', 'Other!B3', '', 'Jump to Other')
	$s.Activate()
	$s.Range('B3').Select()
	$excel.ActiveWindow.FreezePanes = $false
	$excel.ActiveWindow.SplitColumn = 0
	$excel.ActiveWindow.SplitRow = 0
	$s.Range('B3').Select()
	$excel.ActiveWindow.FreezePanes = $true
	$excel.ActiveWindow.DisplayGridlines = $false
	$excel.ActiveWindow.Zoom = 85
	$o = $wb.Worksheets.Item(2)
	$o.Name = 'Other'
	$o.Range('B3').Value2 = 'target'
	$o.Tab.Color = 255
	$s.Activate()
	Save-Xls $wb 'workbook-styles.xls' $null

	# 3. workbook-1904.xls: the 1904 date system.
	$wb = New-Book 1
	$wb.Date1904 = $true
	$d = $wb.Worksheets.Item(1)
	$d.Range('A1').Value2 = 0
	$d.Range('A1').NumberFormat = 'yyyy-mm-dd'
	$d.Range('A2').Value2 = 43830
	$d.Range('A2').NumberFormat = 'yyyy-mm-dd'
	Save-Xls $wb 'workbook-1904.xls' $null

	# 4. workbook-encrypted.xls: password to open (FILEPASS, RC4 CryptoAPI by default).
	$wb = New-Book 1
	$wb.Worksheets.Item(1).Range('A1').Value2 = 'secret'
	Save-Xls $wb 'workbook-encrypted.xls' 'pass'
}
finally {
	$excel.Quit()
	[void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
}
Get-ChildItem $out -Filter *.xls | Select-Object Name, Length
