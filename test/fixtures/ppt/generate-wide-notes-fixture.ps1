param([string]$OutputDirectory=$PSScriptRoot)
# Own pinned synthetic source only; neutral metadata, no objects/macros/links.
$ErrorActionPreference='Stop'
$source=Join-Path $PSScriptRoot 'native-text.ppt'
if((Get-FileHash -LiteralPath $source -Algorithm SHA256).Hash.ToLowerInvariant() -ne '2bae7c1501262b79d2d9840ba968a5ed7b4ee12facde52082d3b114772f9b4e0'){throw 'Expected pinned owned PPT source'}
if(@(Get-Process -Name POWERPNT -ErrorAction SilentlyContinue).Count -gt 0){throw 'PowerPoint already running; coordinate before fixture generation'}
$out=[IO.Path]::GetFullPath($OutputDirectory);[void][IO.Directory]::CreateDirectory($out)
$app=$null;$deck=$null
try {
    $app=New-Object -ComObject PowerPoint.Application
    $app.AutomationSecurity=3;$app.DisplayAlerts=1
    if($app.Presentations.Count -ne 0){throw 'Expected new empty PowerPoint instance'}
    $deck=$app.Presentations.Open($source,-1,0,0)
    $text=([string][char]0x6771+[char]0x4eac+' '+[char]0x03a9+' '+[char]::ConvertFromUtf32(0x1f600)+"`rWide notes.")
    $deck.Slides.Item(2).NotesPage.Shapes.Placeholders.Item(2).TextFrame.TextRange.Text=$text
    $deck.RemoveDocumentInformation(99)
    $props=$deck.BuiltinDocumentProperties
    foreach($entry in @(@('Author','Fixture Author'),@('Title','Owned wide binary PPT notes'))){$property=[System.__ComObject].InvokeMember('Item','GetProperty',$null,$props,@($entry[0]));[void][System.__ComObject].InvokeMember('Value','SetProperty',$null,$property,@($entry[1]))}
    $deck.SaveCopyAs((Join-Path $out 'wide-notes.ppt'),1,0)
    Write-Output (Join-Path $out 'wide-notes.ppt')
} finally {
    try{if($null -ne $deck){$deck.Close();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($deck)}}finally{
        if($null -ne $app){$app.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
        [GC]::Collect();[GC]::WaitForPendingFinalizers()
    }
}
