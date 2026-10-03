param([Parameter(Mandatory)][string]$InputPath,[Parameter(Mandatory)][string]$OutputPath,[string]$SaveCopyPath,[switch]$CaptureBlankCells,[switch]$CaptureCellTypes,[switch]$CaptureWordCharacterFonts,[switch]$CaptureParagraphFormat,[switch]$CapturePptCharacterFonts,[switch]$CapturePptNotes)
$ErrorActionPreference='Stop'
$output=[IO.Path]::GetFullPath($OutputPath)
$arguments=@('-STA','-NoProfile','-File',(Join-Path $PSScriptRoot 'native-office-snapshot.ps1'),'-InputPath',[IO.Path]::GetFullPath($InputPath),'-OutputPath',$output)
if($SaveCopyPath){$arguments+=@('-SaveCopyPath',[IO.Path]::GetFullPath($SaveCopyPath))}
foreach($name in @('CaptureBlankCells','CaptureCellTypes','CaptureWordCharacterFonts','CaptureParagraphFormat','CapturePptCharacterFonts','CapturePptNotes')){if((Get-Variable -Name $name -ValueOnly)){$arguments+='-'+$name}}
$quoted=$arguments|ForEach-Object {'"'+$_.Replace('"','\"')+'"'}
$task=Start-Process -FilePath (Get-Command pwsh).Source -ArgumentList $quoted -WindowStyle Hidden -PassThru
$finished=$task.WaitForExit(45000)
if(-not $finished){Stop-Process -Id $task.Id -ErrorAction SilentlyContinue}
if(Test-Path -LiteralPath ($output+'.process.json')){
    $proof=Get-Content -LiteralPath ($output+'.process.json') -Raw|ConvertFrom-Json
    $owned=Get-Process -Id $proof.processId -ErrorAction SilentlyContinue
    if($null -ne $owned -and $owned.StartTime.ToUniversalTime().Ticks -eq $proof.startTicks -and $owned.Path -eq $proof.executable -and [IO.Path]::GetFileName($owned.Path) -in @('WINWORD.EXE','EXCEL.EXE','POWERPNT.EXE')){Stop-Process -Id $owned.Id}
}
if(-not $finished){throw 'Owned native Office case exceeded45seconds'}
if($task.ExitCode -ne 0){throw "Native Office case failed with exit $($task.ExitCode)"}
Write-Output $output
