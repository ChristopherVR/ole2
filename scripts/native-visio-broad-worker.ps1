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
    return [ordered]@{consumer='Microsoft Visio';version=[string]$app.Version;fullBuild=[string]$app.FullBuild;documentVersion=$document.Version;macroDisabledOpenFlags=458;eventsEnabled=[bool]$app.EventsEnabled;pages=$pages;masters=$masters;styleCount=$document.Styles.Count;fullFidelity=$false;styles=@($document.Styles | ForEach-Object {[ordered]@{id=$_.ID;name=[string]$_.NameU}})}
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
        $page.NameU='Owned hierarchy page'
        $page.PageSheet.CellsU('PageWidth').FormulaU='8 in'
        $page.PageSheet.CellsU('PageHeight').FormulaU='11 in'
        $red=$page.DrawRectangle(1,1,4,3)
        $red.NameU='Overlap red first'
        $red.Text='Owned red'
        $red.CellsU('FillForegnd').FormulaU='RGB(200,20,20)'
        $blue=$page.DrawRectangle(2,2,5,4)
        $blue.NameU='Overlap blue second'
        $blue.Text='Owned blue'
        $blue.CellsU('FillForegnd').FormulaU='RGB(20,20,200)'
        $red.BringToFront()

        $childA=$page.DrawRectangle(1,5,2,6)
        $childA.NameU='Owned child A'
        $childA.Text=('Child '+[char]0x03A9)
        $childB=$page.DrawRectangle(3,5,4,6)
        $childB.NameU='Owned child B'
        $childB.Text='Child B'
        $selection=$page.CreateSelection(0,0,$null)
        $selection.Select($childA,2)
        $selection.Select($childB,2)
        $group=$selection.Group()
        $group.NameU='Owned group'
        $group.Text='Group text'
        $group.CellsU('Angle').FormulaU='15 deg'

        $master=$doc.Masters.Add()
        $master.NameU='Owned rectangle master'
        $masterShape=$master.DrawRectangle(0,0,2,1)
        $masterShape.NameU='Owned master rectangle'
        $masterShape.Text='Master text'
        $instance=$page.Drop($master,6,8)
        $instance.NameU='Owned master instance'
        $instance.Text=('Instance '+[char]0x03A9)
        $instance.CellsU('Char.Size').FormulaU='14 pt'

        $style=$doc.Styles.Add('Owned neutral style','',1,1,1)
        $style.CellsU('Char.Size').FormulaU='18 pt'
        $style.CellsU('LineWeight').FormulaU='0.02 in'
        $style.CellsU('FillForegnd').FormulaU='RGB(40,160,80)'
        $red.Style='Owned neutral style'
        $layer=$page.Layers.Add('Owned foreground')
        $layer.Add($red,0)
        $layer.Add($group,0)

        $page2=$doc.Pages.Add()
        $page2.NameU='Owned geometry page'
        $page2.PageSheet.CellsU('PageWidth').FormulaU='11 in'
        $page2.PageSheet.CellsU('PageHeight').FormulaU='8 in'
        $geometry=$page2.DrawRectangle(1,1,5,3)
        $geometry.NameU='Owned two geometry sections'
        $geometry.Text='Two paths'
        [void]$geometry.AddSection(11)
        [void]$geometry.AddRow(11,1,138)
        [void]$geometry.AddRow(11,2,139)
        $geometry.CellsU('Geometry2.X1').ResultIU=0.0
        $geometry.CellsU('Geometry2.Y1').ResultIU=0.5
        $geometry.CellsU('Geometry2.X2').ResultIU=4.0
        $geometry.CellsU('Geometry2.Y2').ResultIU=1.5
        $geometry.CellsU('Geometry1.NoShow').FormulaU='TRUE'
        $geometry.CellsU('Geometry2.NoFill').FormulaU='TRUE'
        $geometry.CellsU('Geometry2.NoLine').FormulaU='FALSE'
        $line=$page2.DrawLine(1,5,7,6)
        $line.NameU='Owned diagonal line'
        $label=$page2.DrawRectangle(6,1,9,2)
        $label.NameU='Owned Unicode label'
        $label.Text=([string][char]0x6771+[char]0x4EAC+' '+[char]0x03A9+' '+[char]::ConvertFromUtf32(0x1F600))
        $layer2=$page2.Layers.Add('Owned annotations')
        $layer2.Add($label,0)
        $doc.Title='Owned native hierarchy and geometry fixture'
        $doc.Creator='Fixture Author'
        $doc.Company='Fixture Organization'
        $doc.Description='Original neutral pages, group, master, style, layers and geometry; no source documents, macros, objects or links.'
        $doc.Version=0xB0000
        $doc.SaveAsEx([IO.Path]::GetFullPath($SavePath),0)
    } else {
        $inputFile=(Resolve-Path -LiteralPath $InputPath).Path
        if((Get-Item -LiteralPath $inputFile).Length -gt 8388608){throw 'Input exceeds8MiB'}
        $stage='open'
        $doc=$app.Documents.OpenEx($inputFile,458)
        $opened=$true
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
