param(
    [Parameter(Mandatory)][string]$InputPath,
    [Parameter(Mandatory)][string]$OutputPath,
    [int]$MaxCells = 10000,
    [switch]$CaptureRichText,
    [int]$MaxRichTextCharacters = 4096,
    [string]$RichTextCells = '',
    [switch]$CaptureFieldLocations,
    [switch]$CaptureWordCharacterFonts,
    [int]$MaxWordCharacters = 4096,
    [switch]$CaptureCellTypes,
    [switch]$CapturePptCharacterFonts,
    [int]$MaxPptCharacters = 4096,
    [switch]$CaptureBlankCells,
    [switch]$CaptureParagraphFormat,
    [switch]$CapturePptNotes,
    [string]$SaveCopyPath
)
# Explicit fixture paths only. Never enumerate recent documents or run macros.
$ErrorActionPreference = 'Stop'
$inputFile = (Resolve-Path -LiteralPath $InputPath).Path
$app = $null
$document = $null
$snapshot = $null
$appOwned=$false
$processName=switch([IO.Path]::GetExtension($inputFile).ToLowerInvariant()){'.doc'{'WINWORD'}'.xls'{'EXCEL'}'.ppt'{'POWERPNT'}default{throw 'Unsupported file extension'}}
$previousProcessIds=@(Get-Process -Name $processName -ErrorAction SilentlyContinue|ForEach-Object Id)
if($processName -eq 'POWERPNT' -and $previousProcessIds.Count -gt 0){throw 'PowerPoint already running; coordinate before an oracle session'}
function RecordOwnedProcess {
    $newProcesses=@(Get-Process -Name $processName -ErrorAction SilentlyContinue|Where-Object {$previousProcessIds -notcontains $_.Id})
    if($newProcesses.Count -ne 1){throw 'Cannot prove exactly one newly created owned Office process'}
    $owned=$newProcesses[0]
    [ordered]@{processId=$owned.Id;startTicks=$owned.StartTime.ToUniversalTime().Ticks;executable=$owned.Path;ownership='Single new process since explicit COM activation; no preexisting process ID reused'}|ConvertTo-Json|Set-Content ($OutputPath+'.process.json') -Encoding utf8
    $script:appOwned=$true
}
try {
    switch ([IO.Path]::GetExtension($inputFile).ToLowerInvariant()) {
        '.doc' {
            $app = New-Object -ComObject Word.Application
            RecordOwnedProcess
            if($app.Documents.Count -ne 0){throw 'Expected new empty Word instance'}
            $app.Visible = $false
            $app.DisplayAlerts = 0
            $app.AutomationSecurity = 3 # msoAutomationSecurityForceDisable
            # Supplied neutral fixtures contain no external links. Do not change
            # Word Options, which can persist beyond the owned app instance.
            $document = $app.Documents.Open($inputFile, $false, $true, $false, '', '', $false, '', '', 0, 0, $false, $false, 0, $true)
            $paragraphs = @()
            foreach ($paragraph in $document.Paragraphs) {
                $range = $paragraph.Range
                $paragraphEntry = [ordered]@{ text = $range.Text; style = [string]$range.Style.NameLocal; bold = $range.Font.Bold; italic = $range.Font.Italic; size = $range.Font.Size; font = $range.Font.Name }
                if($CaptureParagraphFormat){
                    $format=$range.ParagraphFormat
                    $paragraphEntry['paragraphFormat']=[ordered]@{alignment=$format.Alignment;readingOrder=$format.ReadingOrder;leftIndent=$format.LeftIndent;rightIndent=$format.RightIndent;firstLineIndent=$format.FirstLineIndent;spaceBefore=$format.SpaceBefore;spaceAfter=$format.SpaceAfter;lineSpacing=$format.LineSpacing;lineSpacingRule=$format.LineSpacingRule;keepTogether=$format.KeepTogether;keepWithNext=$format.KeepWithNext;pageBreakBefore=$format.PageBreakBefore;outlineLevel=$format.OutlineLevel;listType=$range.ListFormat.ListType;listString=[string]$range.ListFormat.ListString}
                }
                if ($CaptureWordCharacterFonts) {
                    if ($document.Content.End -gt $MaxWordCharacters) { throw "Word character capture exceeds MaxWordCharacters=$MaxWordCharacters" }
                    $characters = @()
                    foreach ($character in $range.Characters) {
                        $font = $character.Font
                        $characters += [ordered]@{ font = $font.Name; size = $font.Size; bold = $font.Bold; italic = $font.Italic; underline = $font.Underline; color = $font.Color }
                    }
                    $paragraphEntry['characterFonts'] = $characters
                }
                $paragraphs += $paragraphEntry
            }
            $headers = @()
            foreach ($section in $document.Sections) {
                foreach ($header in $section.Headers) { if ($header.Exists) { $headers += [ordered]@{ section = $section.Index; kind = $header.Index; text = $header.Range.Text } } }
                foreach ($footer in $section.Footers) { if ($footer.Exists) { $headers += [ordered]@{ section = $section.Index; kind = -$footer.Index; text = $footer.Range.Text } } }
            }
            $fields = @()
            foreach ($field in $document.Fields) {
                $fieldEntry = [ordered]@{ type = $field.Type; code = $field.Code.Text; result = $field.Result.Text }
                if ($CaptureFieldLocations) {
                    $position = $field.Code.Start
                    $paragraphIndex = 0
                    foreach ($paragraph in $document.Paragraphs) {
                        if ($position -ge $paragraph.Range.Start -and $position -lt $paragraph.Range.End) {
                            $fieldEntry['paragraphIndex'] = $paragraphIndex
                            $fieldEntry['relativeCodeStart'] = $position - $paragraph.Range.Start
                            break
                        }
                        $paragraphIndex++
                    }
                    if (-not $fieldEntry.Contains('paragraphIndex')) { throw 'Main-story field location not found' }
                }
                $fields += $fieldEntry
            }
            $footnotes = @($document.Footnotes | ForEach-Object { $_.Range.Text })
            $snapshot = [ordered]@{ consumer = 'Microsoft Word'; version = $app.Version; paragraphs = $paragraphs; tables = $document.Tables.Count; shapes = $document.Shapes.Count; inlineShapes = $document.InlineShapes.Count; sections = $document.Sections.Count; headers = $headers; fields = $fields; footnotes = $footnotes }
        }
        '.xls' {
            $app = New-Object -ComObject Excel.Application
            RecordOwnedProcess
            if($app.Workbooks.Count -ne 0){throw 'Expected new empty Excel instance'}
            $app.Visible = $false
            $app.DisplayAlerts = $false
            $app.AutomationSecurity = 3
            $app.EnableEvents = $false
            # Explicit empty passwords prevent password dialogs. UpdateLinks=0.
            $openArguments=[object[]]@($inputFile,0,$true,[Type]::Missing,'','',$true,[Type]::Missing,[Type]::Missing,$false,$false,[Type]::Missing,$false,$false,0)
            # Dispatch directly: pwsh's reflected15-argument overload rejects
            # Missing even though IDispatch optional VARIANT parameters allow it.
            $document=[System.__ComObject].InvokeMember('Open','InvokeMethod',$null,$app.Workbooks,$openArguments) # xlNormalLoad, no recovery
            $app.Calculation = -4135 # manual; do not recalculate during snapshot
            $sheets = @()
            $richSelection = @($RichTextCells.Split(';', [StringSplitOptions]::RemoveEmptyEntries))
            $richSeen = @{}
            foreach ($sheet in $document.Worksheets) {
                $range = $sheet.UsedRange
                if ([long]$range.Rows.Count * $range.Columns.Count -gt $MaxCells) { throw "Used range exceeds MaxCells=$MaxCells" }
                $cells = @()
                foreach ($cell in $range.Cells) {
                    if ($CaptureBlankCells -or $null -ne $cell.Value2 -or $cell.HasFormula) {
                        $entry = [ordered]@{ row = $cell.Row - 1; col = $cell.Column - 1; value = $cell.Value2; formula = $cell.Formula; numberFormat = $cell.NumberFormat; bold = $cell.Font.Bold; italic = $cell.Font.Italic; font = $cell.Font.Name; size = $cell.Font.Size; merge = [string]$cell.MergeArea.Address() }
                        if($CaptureBlankCells){$entry['hasFormula']=[bool]$cell.HasFormula;$entry['rowHeight']=$cell.RowHeight;$entry['columnWidth']=$cell.ColumnWidth;$entry['fillColor']=$cell.Interior.Color;$entry['fontColor']=$cell.Font.Color;$entry['wrapText']=$cell.WrapText;$entry['horizontalAlignment']=$cell.HorizontalAlignment}
                        if ($CaptureCellTypes) {
                            # WorksheetFunction.IsError(range) can return a COM array for
                            # merged cells. A scalar local reference avoids array coercion.
                            $entry['isError'] = [bool]$sheet.Evaluate('ISERROR(' + $cell.Address() + ')')
                            $entry['displayText'] = [string]$cell.Text
                            $entry['valueType'] = if ($null -eq $cell.Value2) { 'null' } else { $cell.Value2.GetType().FullName }
                        }
                        $richKey = $sheet.Name + '!' + ($cell.Row - 1) + ',' + ($cell.Column - 1)
                        if ($CaptureRichText -and ($richSelection.Count -eq 0 -or $richSelection -contains $richKey) -and $cell.Value2 -is [string]) {
                            $richSeen[$richKey] = $true
                            $text = [string]$cell.Value2
                            if ($text.Length -gt $MaxRichTextCharacters) { throw "Rich text exceeds MaxRichTextCharacters=$MaxRichTextCharacters" }
                            $characters = @()
                            for ($index = 1; $index -le $text.Length; $index++) {
                                $font = $cell.Characters($index, 1).Font
                                $characters += [ordered]@{ font = $font.Name; size = $font.Size; bold = $font.Bold; italic = $font.Italic; underline = $font.Underline; color = $font.Color }
                            }
                            $entry['characterFonts'] = $characters
                        }
                        $cells += $entry
                    }
                }
                $sheets += [ordered]@{ name = $sheet.Name; visible = $sheet.Visible; cells = $cells; shapes = $sheet.Shapes.Count; comments = $sheet.Comments.Count; hyperlinks = $sheet.Hyperlinks.Count }
            }
            if ($CaptureRichText) {
                foreach ($selected in $richSelection) {
                    if (-not $richSeen.ContainsKey($selected)) { throw "Requested rich text cell missing or non-string: $selected" }
                }
            }
            $repairMode=$document.RepairMode
            $snapshot = [ordered]@{ consumer = 'Microsoft Excel'; version = $app.Version; corruptLoad = 0; repairMode = if($null -eq $repairMode){$null}else{[bool]$repairMode}; date1904 = $document.Date1904; sheets = $sheets; names = @($document.Names | ForEach-Object { [ordered]@{ name = $_.Name; refersTo = $_.RefersTo } }) }
        }
        '.ppt' {
            $app = New-Object -ComObject PowerPoint.Application
            RecordOwnedProcess
            if($app.Presentations.Count -ne 0){throw 'Expected new empty PowerPoint instance'}
            $app.AutomationSecurity = 3
            $app.DisplayAlerts = 1 # ppAlertsNone
            $document = $app.Presentations.Open($inputFile, -1, 0, 0)
            $slides = @()
            $pptCharacterCount = 0
            foreach ($slide in $document.Slides) {
                $shapes = @()
                foreach ($shape in $slide.Shapes) {
                    $text = if ($shape.HasTextFrame -and $shape.TextFrame.HasText) { $shape.TextFrame.TextRange.Text } else { '' }
                    $shapeEntry = [ordered]@{ name = $shape.Name; type = $shape.Type; text = $text; left = $shape.Left; top = $shape.Top; width = $shape.Width; height = $shape.Height }
                    if ($CapturePptCharacterFonts) {
                        $pptCharacterCount += $text.Length
                        if ($pptCharacterCount -gt $MaxPptCharacters) { throw "PPT character capture exceeds MaxPptCharacters=$MaxPptCharacters" }
                        $characters = @()
                        for ($index = 1; $index -le $text.Length; $index++) {
                            $font = $shape.TextFrame.TextRange.Characters($index, 1).Font
                            $characters += [ordered]@{ font = $font.Name; size = $font.Size; bold = $font.Bold; italic = $font.Italic; underline = $font.Underline; color = $font.Color.RGB }
                        }
                        $shapeEntry['characterFonts'] = $characters
                    }
                    $shapes += $shapeEntry
                }
                $notes = @()
                $noteShapes = @()
                foreach ($shape in $slide.NotesPage.Shapes) {
                    if ($shape.HasTextFrame -and $shape.TextFrame.HasText) { $notes += $shape.TextFrame.TextRange.Text }
                    if($CapturePptNotes){
                        $noteText=if($shape.HasTextFrame -and $shape.TextFrame.HasText){[string]$shape.TextFrame.TextRange.Text}else{''}
                        $noteEntry=[ordered]@{id=$shape.Id;name=[string]$shape.Name;type=$shape.Type;text=$noteText;left=$shape.Left;top=$shape.Top;width=$shape.Width;height=$shape.Height}
                        $fonts=@()
                        $pptCharacterCount+=$noteText.Length
                        if($pptCharacterCount -gt $MaxPptCharacters){throw 'PPT note character capture exceeds bound'}
                        for($index=1;$index -le $noteText.Length;$index++){$font=$shape.TextFrame.TextRange.Characters($index,1).Font;$fonts+=[ordered]@{font=$font.Name;size=$font.Size;bold=$font.Bold;italic=$font.Italic;underline=$font.Underline;color=$font.Color.RGB}}
                        $noteEntry['characterFonts']=$fonts
                        $noteShapes+=$noteEntry
                    }
                }
                $slideEntry=[ordered]@{ id = $slide.SlideID; shapes = $shapes; notes = $notes }
                if($CapturePptNotes){$slideEntry['noteShapes']=$noteShapes}
                $slides += $slideEntry
            }
            $snapshot = [ordered]@{ consumer = 'Microsoft PowerPoint'; version = $app.Version; width = $document.PageSetup.SlideWidth; height = $document.PageSetup.SlideHeight; slides = $slides }
        }
        default { throw 'Supported explicit inputs: .doc, .xls, .ppt' }
    }
    $snapshot | ConvertTo-Json -Depth 30 | Set-Content -LiteralPath $OutputPath -Encoding utf8
    if($SaveCopyPath){
        $copy=[IO.Path]::GetFullPath($SaveCopyPath)
        if($copy -eq $inputFile){throw 'SaveCopyPath must differ from the source'}
        switch([IO.Path]::GetExtension($inputFile).ToLowerInvariant()){
            '.xls' {$document.SaveCopyAs($copy)}
            '.ppt' {$document.SaveCopyAs($copy,1,0)}
            '.doc' {$document.RemovePersonalInformation=$true;$document.SaveAs2($copy,0,$false,'',$false)}
        }
    }
}
finally {
    if ($null -ne $document) {
        if ([IO.Path]::GetExtension($inputFile) -eq '.ppt') { $document.Close() } else { $document.Close($false) }
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($document)
    }
    if ($null -ne $app) {
        if($appOwned){$app.Quit()}
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
