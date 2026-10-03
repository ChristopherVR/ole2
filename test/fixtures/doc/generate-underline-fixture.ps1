param([string]$OutputDirectory=$PSScriptRoot)
# Derive only from pinned owned synthetic content, never personal documents.
$ErrorActionPreference='Stop'
$source=Join-Path $PSScriptRoot 'rich-runs.doc'
if((Get-FileHash -LiteralPath $source -Algorithm SHA256).Hash.ToLowerInvariant() -ne '4b1cdf12f247fc235fdbee7232c3ed1d16d6380a98810b709b05e516cbfdd72f'){throw 'Expected pinned owned rich fixture'}
$out=[IO.Path]::GetFullPath($OutputDirectory);[void][IO.Directory]::CreateDirectory($out)
$app=$null;$doc=$null
try {
    $app=New-Object -ComObject Word.Application
    $app.Visible=$false;$app.DisplayAlerts=0;$app.AutomationSecurity=3
    # Pinned neutral fixture has no links/fields/objects. No global Options change.
    $doc=$app.Documents.Open($source,$false,$true,$false,'','',$false,'','',0,0,$false,$false,0,$true)
    $underlineStart=$doc.Paragraphs.Item(2).Range.Start; $doc.Range($underlineStart,$underlineStart+9).Font.Underline=1 # wdUnderlineSingle on Bold text only
    $doc.RemovePersonalInformation=$true;$doc.EmbedTrueTypeFonts=$false
    $doc.SaveAs2((Join-Path $out 'underline-runs.doc'),0,$false,'',$false)
    Write-Output (Join-Path $out 'underline-runs.doc')
} finally {
    try{if($null -ne $doc){$doc.Close($false);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($doc)}}finally{
        if($null -ne $app){$app.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
        [GC]::Collect();[GC]::WaitForPendingFinalizers()
    }
}

