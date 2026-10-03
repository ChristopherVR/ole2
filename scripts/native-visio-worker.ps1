param([Parameter(Mandatory)][ValidateSet('snapshot','generate','save-copy')][string]$Action,[string]$InputPath,[Parameter(Mandatory)][string]$OutputPath,[string]$SavePath)
$ErrorActionPreference='Stop'
$app=$null
$doc=$null
$owned=$null
$opened=$false
$stage='activation'
Add-Type -TypeDefinition @'
using System;using System.Runtime.InteropServices;
public static class VisioOwnedWindow { [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hwnd,out uint pid); }
'@
function CellSnapshot($shape,$name) {
    if (-not $shape.CellExistsU($name,0)) { return $null }
    $cell=$shape.CellsU($name)
    return [ordered]@{formula=[string]$cell.FormulaU;result=[double]$cell.ResultIU}
}
function ShapeSnapshot($shape,$depth) {
    if($depth -gt 32){throw 'Shape depth exceeds32'}
    $script:shapeCount++
    if($script:shapeCount -gt 500){throw 'Shape count exceeds500'}
    $cells=[ordered]@{}
    foreach($name in @('PinX','PinY','Width','Height','LocPinX','LocPinY','Angle','FlipX','FlipY','TxtPinX','TxtPinY','TxtWidth','TxtHeight','TxtLocPinX','TxtLocPinY','TxtAngle','LineColor','LinePattern','LineWeight','FillForegnd','FillPattern','Char.Size','Char.Font','Para.HorzAlign','LayerMember')) {$cells[$name]=CellSnapshot $shape $name}
    $geometry=@()
    for($section=10;$section -lt 74;$section++) {
        if(-not $shape.SectionExists($section,0)){continue}
        $rows=$shape.RowCount($section)
        if($rows -gt 4096){throw 'Geometry row count exceeds4096'}
        $entries=@()
        for($row=0;$row -lt $rows;$row++) {
            $values=@()
            for($column=0;$column -lt 5;$column++) {
                if($shape.CellsSRCExists($section,$row,$column,0)) {
                    $cell=$shape.CellsSRC($section,$row,$column)
                    $values += [ordered]@{column=$column;formula=[string]$cell.FormulaU;result=[double]$cell.ResultIU}
                }
            }
            $entries += [ordered]@{row=$row;type=$shape.RowType($section,$row);cells=$values}
        }
        $geometry += [ordered]@{section=$section;rows=$entries}
    }
    $master=$shape.Master
    $children=@()
    if($shape.Type -eq 2){foreach($child in $shape.Shapes){$children+=ShapeSnapshot $child ($depth+1)}}
    return [ordered]@{id=$shape.ID;name=[string]$shape.NameU;type=$shape.Type;text=[string]$shape.Text;masterId=if($null -eq $master){$null}else{$master.ID};fields=$shape.Fields.Count;hyperlinks=$shape.Hyperlinks.Count;cells=$cells;geometry=$geometry;children=$children}
}
function DocumentSnapshot($document) {
    $script:shapeCount=0
    $pages=@()
    if($document.Pages.Count -gt 100 -or $document.Masters.Count -gt 100){throw 'Page/master count exceeds100'}
    foreach($page in $document.Pages){
        $pageCells=[ordered]@{}
        foreach($name in @('PageWidth','PageHeight','PageScale','DrawingScale')){$pageCells[$name]=CellSnapshot $page.PageSheet $name}
        $shapes=@();foreach($shape in $page.Shapes){$shapes+=ShapeSnapshot $shape 0}
        $layers=@();foreach($layer in $page.Layers){$layers += [ordered]@{index=$layer.Index;name=[string]$layer.NameU}}
        $pages += [ordered]@{id=$page.ID;name=[string]$page.NameU;background=[bool]$page.Background;cells=$pageCells;layers=$layers;shapes=$shapes;connects=$page.Connects.Count}
    }
    $masters=@();foreach($master in $document.Masters){$shapes=@();foreach($shape in $master.Shapes){$shapes+=ShapeSnapshot $shape 0};$masters += [ordered]@{id=$master.ID;name=[string]$master.NameU;shapes=$shapes}}
    return [ordered]@{consumer='Microsoft Visio';version=[string]$app.Version;fullBuild=[string]$app.FullBuild;documentVersion=$document.Version;macroDisabledOpenFlags=458;eventsEnabled=[bool]$app.EventsEnabled;pages=$pages;masters=$masters;styleCount=$document.Styles.Count;fullFidelity=$false}
}
try {
    $app=New-Object -ComObject Visio.InvisibleApp
    $app.Visible=$false
    $app.EventsEnabled=$false
    $app.AlertResponse=7
    if($app.Documents.Count -ne 0){throw 'Expected new empty Visio instance'}
    # Visio.ProcessID is an application identity, not the Windows PID.
    # Resolve only this owned invisible application's exact window handle.
    [uint32]$windowsPid=0
    [void][VisioOwnedWindow]::GetWindowThreadProcessId([IntPtr]$app.WindowHandle32,[ref]$windowsPid)
    $process=Get-Process -Id $windowsPid -ErrorAction SilentlyContinue
    if($null -ne $process){
        $owned=[ordered]@{processId=$process.Id;startTicks=$process.StartTime.ToUniversalTime().Ticks;executable=$process.Path}
        $owned|ConvertTo-Json|Set-Content ($OutputPath+'.process.json') -Encoding utf8
    }
    if($Action -eq 'generate'){
        $stage='generate'
        $doc=$app.Documents.Add('')
        $page=$doc.Pages.Item(1)
        $page.NameU='Owned page'
        $page.PageSheet.CellsU('PageWidth').FormulaU='8 in'
        $page.PageSheet.CellsU('PageHeight').FormulaU='11 in'
        $shape=$page.DrawRectangle(1,2,5,4)
        $shape.NameU='Owned rectangle'
        $shape.Text="Hello`n"
        $shape.CellsU('Char.Size').FormulaU='12 pt'
        $shape.CellsU('Char.Font').FormulaU='0'
        $shape.CellsU('FillPattern').FormulaU='0'
        $shape.CellsU('PinX').FormulaU='3 in'
        $shape.CellsU('PinY').FormulaU='3 in'
        $line=$page.DrawLine(1,6,5,6)
        $line.NameU='Owned line'
        $label=$page.DrawRectangle(1,8,5,9)
        $label.NameU='Owned label'
        $label.Text=('Native owned label '+[char]0x03A9)
        $label.CellsU('LinePattern').FormulaU='0'
        $label.CellsU('FillPattern').FormulaU='0'
        $label.CellsU('Char.Size').FormulaU='12 pt'
        $doc.Title='Owned native legacy Visio fixture'
        $doc.Creator='Fixture Author'
        $doc.Company='Fixture Organization'
        $doc.Description='Synthetic rectangle only; no objects, macros, external data or source documents.'
        $doc.Version=0xB0000
        $doc.SaveAsEx([IO.Path]::GetFullPath($SavePath),0)
    } else {
        $inputFile=(Resolve-Path -LiteralPath $InputPath).Path
        if((Get-Item -LiteralPath $inputFile).Length -gt 8388608){throw 'Input exceeds8MiB'}
        $stage='open'
        $doc=$app.Documents.OpenEx($inputFile,458)
        $opened=$true
        [ordered]@{accepted=$true;version=[string]$app.Version;pages=$doc.Pages.Count;shapes=$doc.Pages.Item(1).Shapes.Count;macrosDisabled=$true;eventsEnabled=[bool]$app.EventsEnabled}|ConvertTo-Json|Set-Content ($OutputPath+'.open.json') -Encoding utf8
        if($Action -eq 'save-copy'){$doc.Version=0xB0000;$doc.SaveAsEx([IO.Path]::GetFullPath($SavePath),0)}
    }
    $stage='snapshot'
    DocumentSnapshot $doc|ConvertTo-Json -Depth 50|Set-Content -LiteralPath $OutputPath -Encoding utf8
} catch {
    [ordered]@{consumer='Microsoft Visio';version=if($null -eq $app){$null}else{[string]$app.Version};action=$Action;stage=$stage;accepted=$opened;message=$_.Exception.Message;hresult=$_.Exception.HResult;stack=$_.ScriptStackTrace}|ConvertTo-Json -Depth 10|Set-Content -LiteralPath $OutputPath -Encoding utf8
    exit 1
} finally {
    try{if($null -ne $doc){$doc.Saved=$true;$doc.Close();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($doc)}}finally{
        if($null -ne $app){$app.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
        [GC]::Collect();[GC]::WaitForPendingFinalizers()
    }
}
