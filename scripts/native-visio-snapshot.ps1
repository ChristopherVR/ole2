param(
    [ValidateSet('snapshot','generate','generate-literal','save-copy')][string]$Action='snapshot',
    [string]$InputPath,
    [Parameter(Mandatory)][string]$OutputPath,
    [string]$SavePath
)
# New owned invisible Visio instances only. Never attaches to user sessions.
# Readonly + DontList + Hidden + MacrosDisabled + NoWorkspace =458.
# https://learn.microsoft.com/en-us/office/vba/api/visio.documents.openex
# EventsEnabled=false also prevents RUNADDON execution on formula evaluation.
# https://learn.microsoft.com/en-us/office/vba/api/visio.application.eventsenabled
$ErrorActionPreference='Stop'
if($Action -notin @('generate','generate-literal') -and -not $InputPath){throw 'InputPath is required'}
if($Action -ne 'snapshot' -and -not $SavePath){throw 'SavePath is required'}
$output=[IO.Path]::GetFullPath($OutputPath)
$worker=Join-Path $PSScriptRoot 'native-visio-worker.ps1'
$arguments=@('-STA','-NoProfile','-File',$worker,'-Action',$Action,'-OutputPath',$output)
if($InputPath){$arguments+=@('-InputPath',[IO.Path]::GetFullPath($InputPath))}
if($SavePath){$arguments+=@('-SavePath',[IO.Path]::GetFullPath($SavePath))}
# Start-Process joins its argument array: quote each argument for Windows argv.
# No shell, expression evaluation, or document content is interpolated here.
$quoted=$arguments|ForEach-Object {'"'+$_.Replace('"','\"')+'"'}
$task=Start-Process -FilePath (Get-Command pwsh).Source -ArgumentList $quoted -WindowStyle Hidden -PassThru
$finished=$task.WaitForExit(45000)
if(-not $finished){Stop-Process -Id $task.Id -ErrorAction SilentlyContinue}
# Resolve only the exact HWND-derived Windows PID, start time, and executable
# recorded by the new owned app. Visio.ProcessID is NOT the Windows PID.
$proofPath=$output+'.process.json'
if(Test-Path -LiteralPath $proofPath){
    $proof=Get-Content -LiteralPath $proofPath -Raw|ConvertFrom-Json
    $owned=Get-Process -Id $proof.processId -ErrorAction SilentlyContinue
    if($null -ne $owned -and $owned.StartTime.ToUniversalTime().Ticks -eq $proof.startTicks -and $owned.Path -eq $proof.executable -and [IO.Path]::GetFileName($owned.Path) -ieq 'VISIO.EXE'){Stop-Process -Id $owned.Id}
}
if(-not $finished){throw 'Owned native case exceeded45-second runtime bound'}
if(Test-Path -LiteralPath $output){Get-Content -LiteralPath $output -Raw}
if($task.ExitCode -ne 0){throw "Native case failed with helper exit code $($task.ExitCode); see $output"}
