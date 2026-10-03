param([string]$OutputPath=(Join-Path $PSScriptRoot 'native-literal-transform.vsd'),[string]$SnapshotPath=(Join-Path $PSScriptRoot 'native-literal-transform.generated.json'))
# Fresh original rectangle; explicit ResultIU constants clear all XForm formulas.
# Existing native-visio16-v11.vsd is neither read nor modified.
$ErrorActionPreference='Stop'
& (Join-Path $PSScriptRoot '../../../scripts/native-visio-snapshot.ps1') -Action generate-literal -OutputPath $SnapshotPath -SavePath $OutputPath
Get-FileHash -LiteralPath $OutputPath -Algorithm SHA256
