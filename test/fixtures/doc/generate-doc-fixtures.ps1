$ErrorActionPreference='Stop'
$app=$null;$doc=$null
try {
$app=New-Object -ComObject Word.Application
$app.Visible=$false;$app.DisplayAlerts=0;$app.AutomationSecurity=3
$doc=$app.Documents.Add()
$doc.Content.Text="First paragraph plain text.`rSecond paragraph has a field: `r"
$header=$doc.Sections.Item(1).Headers.Item(1).Range
$header.Text='Fixture header '
$fieldRange=$header.Duplicate;$fieldRange.Collapse(0)
[void]$doc.Fields.Add($fieldRange,33)
$range=$doc.Paragraphs.Item(2).Range.Duplicate;$range.End=$range.End-1;$range.Collapse(0)
[void]$doc.Fields.Add($range,-1,'QUOTE "Fixture field"',$false)
$props=$doc.BuiltinDocumentProperties
$prop=[System.__ComObject].InvokeMember('Item','GetProperty',$null,$props,@('Author'))
[void][System.__ComObject].InvokeMember('Value','SetProperty',$null,$prop,@('Fixture Author'))
$doc.RemovePersonalInformation=$true
$doc.SaveAs2((Join-Path $PSScriptRoot 'header-field.doc'),0)
} finally {if($doc){$doc.Close($false);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($doc)};if($app){$app.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}}

