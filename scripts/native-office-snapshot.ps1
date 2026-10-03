param(
    [Parameter(Mandatory)][string]$InputPath,
    [Parameter(Mandatory)][string]$OutputPath,
    [int]$MaxCells = 10000,
    [switch]$CaptureRichText,
    [int]$MaxRichTextCharacters = 4096,
    [string]$RichTextCells = '',
    [switch]$CaptureFieldLocations
)
# Explicit fixture paths only. Never enumerate recent documents or run macros.
$ErrorActionPreference = 'Stop'
$inputFile = (Resolve-Path -LiteralPath $InputPath).Path
$app = $null
$document = $null
$snapshot = $null
try {
    switch ([IO.Path]::GetExtension($inputFile).ToLowerInvariant()) {
        '.doc' {
            $app = New-Object -ComObject Word.Application
            $app.Visible = $false
            $app.DisplayAlerts = 0
            $app.AutomationSecurity = 3 # msoAutomationSecurityForceDisable
            $oldLinks = $app.Options.UpdateLinksAtOpen
            $app.Options.UpdateLinksAtOpen = $false
            $document = $app.Documents.Open($inputFile, $false, $true, $false, '', '', $false, '', '', 0, 0, $false, $false, 0, $true)
            $paragraphs = @()
            foreach ($paragraph in $document.Paragraphs) {
                $range = $paragraph.Range
                $paragraphs += [ordered]@{ text = $range.Text; style = [string]$range.Style.NameLocal; bold = $range.Font.Bold; italic = $range.Font.Italic; size = $range.Font.Size; font = $range.Font.Name }
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
            $app.Visible = $false
            $app.DisplayAlerts = $false
            $app.AutomationSecurity = 3
            $app.EnableEvents = $false
            $app.AskToUpdateLinks = $false
            # Explicit empty passwords prevent password dialogs. UpdateLinks=0.
            $document = $app.Workbooks.Open($inputFile, 0, $true, [Type]::Missing, '', '', $true, [Type]::Missing, [Type]::Missing, $false, $false, [Type]::Missing, $false)
            $app.Calculation = -4135 # manual; do not recalculate during snapshot
            $sheets = @()
            $richSelection = @($RichTextCells.Split(';', [StringSplitOptions]::RemoveEmptyEntries))
            $richSeen = @{}
            foreach ($sheet in $document.Worksheets) {
                $range = $sheet.UsedRange
                if ([long]$range.Rows.Count * $range.Columns.Count -gt $MaxCells) { throw "Used range exceeds MaxCells=$MaxCells" }
                $cells = @()
                foreach ($cell in $range.Cells) {
                    if ($null -ne $cell.Value2 -or $cell.HasFormula) {
                        $entry = [ordered]@{ row = $cell.Row - 1; col = $cell.Column - 1; value = $cell.Value2; formula = $cell.Formula; numberFormat = $cell.NumberFormat; bold = $cell.Font.Bold; italic = $cell.Font.Italic; font = $cell.Font.Name; size = $cell.Font.Size; merge = [string]$cell.MergeArea.Address() }
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
            $snapshot = [ordered]@{ consumer = 'Microsoft Excel'; version = $app.Version; date1904 = $document.Date1904; sheets = $sheets; names = @($document.Names | ForEach-Object { [ordered]@{ name = $_.Name; refersTo = $_.RefersTo } }) }
        }
        '.ppt' {
            $app = New-Object -ComObject PowerPoint.Application
            $app.AutomationSecurity = 3
            $app.DisplayAlerts = 1 # ppAlertsNone
            $document = $app.Presentations.Open($inputFile, -1, 0, 0)
            $slides = @()
            foreach ($slide in $document.Slides) {
                $shapes = @()
                foreach ($shape in $slide.Shapes) {
                    $text = if ($shape.HasTextFrame -and $shape.TextFrame.HasText) { $shape.TextFrame.TextRange.Text } else { '' }
                    $shapes += [ordered]@{ name = $shape.Name; type = $shape.Type; text = $text; left = $shape.Left; top = $shape.Top; width = $shape.Width; height = $shape.Height }
                }
                $notes = @()
                foreach ($shape in $slide.NotesPage.Shapes) {
                    if ($shape.HasTextFrame -and $shape.TextFrame.HasText) { $notes += $shape.TextFrame.TextRange.Text }
                }
                $slides += [ordered]@{ id = $slide.SlideID; shapes = $shapes; notes = $notes }
            }
            $snapshot = [ordered]@{ consumer = 'Microsoft PowerPoint'; version = $app.Version; width = $document.PageSetup.SlideWidth; height = $document.PageSetup.SlideHeight; slides = $slides }
        }
        default { throw 'Supported explicit inputs: .doc, .xls, .ppt' }
    }
    $snapshot | ConvertTo-Json -Depth 30 | Set-Content -LiteralPath $OutputPath -Encoding utf8
}
finally {
    if ($null -ne $document) {
        if ([IO.Path]::GetExtension($inputFile) -eq '.ppt') { $document.Close() } else { $document.Close($false) }
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($document)
    }
    if ($null -ne $app) {
        if ($null -ne $oldLinks) { $app.Options.UpdateLinksAtOpen = $oldLinks }
        $app.Quit()
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
