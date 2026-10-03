param(
    [string]$OutputPath=(Join-Path $PSScriptRoot 'native-visio16-v11.vsd'),
    [string]$SnapshotPath=(Join-Path $PSScriptRoot 'native-visio16-v11.generated.json')
)
# Original synthetic content only: no input documents, templates, stencils,
# embedded objects, macros, external links, or field updates.
# Requires installed Visio. Binary timestamps may vary between generation runs.
$ErrorActionPreference='Stop'
$runner=Join-Path $PSScriptRoot '../../../scripts/native-visio-snapshot.ps1'
& $runner -Action generate -OutputPath $SnapshotPath -SavePath $OutputPath
Get-FileHash -LiteralPath $OutputPath -Algorithm SHA256
