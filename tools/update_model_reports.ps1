<#
.SYNOPSIS
Populate embedded-model report layouts using an installed desktop Excel.
.DESCRIPTION
Works on a private copy, updates model PivotTable layouts, saves, and verifies
the saved report counts after reopening. Never overwrites an existing output.
Does not call RefreshAll or refresh upstream source connections. Macros and
link updates are disabled. This optional tool is not a headless calculation engine.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)][string]$Source,
    [Parameter(Mandatory=$true)][string]$Output
)
$ErrorActionPreference = 'Stop'

function Release-Com($value) {
    if ($null -ne $value -and [Runtime.InteropServices.Marshal]::IsComObject($value)) {
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($value)
    }
}

function Test-ModelCache($cache) {
    $connection=$null
    try {
        if (-not $cache.OLAP) { return $false }
        $connection=$cache.WorkbookConnection
        return $connection.Name -eq 'ThisWorkbookDataModel'
    } finally { Release-Com $connection }
}

function Read-ReportState($workbook) {
    $pivots = @(); $charts = @(); $sheets = $null
    try {
        $sheets = $workbook.Worksheets
        for ($i=1; $i -le $sheets.Count; $i++) {
            $sheet=$null; $tables=$null; $objects=$null
            try {
                $sheet=$sheets.Item($i); $tables=$sheet.PivotTables()
                for ($j=1; $j -le $tables.Count; $j++) {
                    $pivot=$null; $cache=$null; $range=$null
                    try {
                        $pivot=$tables.Item($j); $cache=$pivot.PivotCache()
                        if (-not (Test-ModelCache $cache)) { continue }
                        $range=$pivot.TableRange1; $values=$range.Value2; $populated=0
                        foreach ($value in $values) {
                            if ($null -ne $value -and [string]$value -ne '') { $populated++ }
                        }
                        $pivots += [ordered]@{sheet=$sheet.Name;name=$pivot.Name;populated_cells=$populated}
                    } finally { Release-Com $range; Release-Com $cache; Release-Com $pivot }
                }
                $objects=$sheet.ChartObjects()
                for ($j=1; $j -le $objects.Count; $j++) {
                    $object=$null; $chart=$null; $layout=$null; $pivot=$null; $series=$null; $cache=$null
                    try {
                        $object=$objects.Item($j); $chart=$object.Chart
                        try { $layout=$chart.PivotLayout } catch { continue }
                        if ($null -eq $layout) { continue }
                        $pivot=$layout.PivotTable; $cache=$pivot.PivotCache()
                        if (-not (Test-ModelCache $cache)) { continue }
                        $series=$chart.SeriesCollection()
                        $charts += [ordered]@{sheet=$sheet.Name;name=$object.Name;pivot=$pivot.Name;series=$series.Count}
                    } finally {
                        Release-Com $series; Release-Com $cache; Release-Com $pivot; Release-Com $layout
                        Release-Com $chart; Release-Com $object
                    }
                }
            } finally { Release-Com $objects; Release-Com $tables; Release-Com $sheet }
        }
    } finally { Release-Com $sheets }
    return [ordered]@{pivot_count=$pivots.Count;chart_count=$charts.Count;pivots=$pivots;charts=$charts}
}

$sourcePath=(Resolve-Path -LiteralPath $Source).Path
$outputPath=$ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Output)
if ([IO.Path]::GetExtension($sourcePath) -ne '.xlsx' -or [IO.Path]::GetExtension($outputPath) -ne '.xlsx') {
    throw 'Source and Output must be .xlsx workbooks.'
}
if (Test-Path -LiteralPath $outputPath) { throw 'Output already exists; choose a new filename.' }
$parent=[IO.Path]::GetDirectoryName($outputPath)
if (-not (Test-Path -LiteralPath $parent -PathType Container)) { throw 'Output directory does not exist.' }
$stageDirectory=Join-Path $parent ('.excelmodel-update-'+[Guid]::NewGuid().ToString('N'))
$stageFile=Join-Path $stageDirectory ([IO.Path]::GetFileName($outputPath))
$excel=$null; $books=$null; $book=$null; $model=$null
try {
    [void](New-Item -ItemType Directory -Path $stageDirectory)
    Copy-Item -LiteralPath $sourcePath -Destination $stageFile
    $excel=New-Object -ComObject Excel.Application
    $excel.Visible=$false; $excel.DisplayAlerts=$false
    $excel.AutomationSecurity=3; $excel.EnableEvents=$false
    $books=$excel.Workbooks
    $book=$books.Open($stageFile,0,$false)
    $model=$book.Model; $model.Initialize(); Release-Com $model; $model=$null
    $before=Read-ReportState $book
    if ($before.pivot_count -eq 0) { throw 'No embedded-model PivotTables were found.' }
    $sheets=$null
    try {
        $sheets=$book.Worksheets
        for ($i=1; $i -le $sheets.Count; $i++) {
            $sheet=$null; $pivots=$null
            try {
                $sheet=$sheets.Item($i); $pivots=$sheet.PivotTables()
                for ($j=1; $j -le $pivots.Count; $j++) {
                    $pivot=$null; $cache=$null
                    try {
                        $pivot=$pivots.Item($j); $cache=$pivot.PivotCache()
                        if (Test-ModelCache $cache) { $pivot.Update() }
                    } finally { Release-Com $cache; Release-Com $pivot }
                }
            } finally { Release-Com $pivots; Release-Com $sheet }
        }
    } finally { Release-Com $sheets }
    $updated=Read-ReportState $book
    if ($updated.pivot_count -ne $before.pivot_count -or $updated.chart_count -ne $before.chart_count) {
        throw 'Report counts changed while updating layouts.'
    }
    if (@($updated.pivots | Where-Object { $_.populated_cells -eq 0 }).Count -gt 0) {
        throw 'At least one PivotTable still has no displayed cells after Update.'
    }
    $book.Save(); $book.Close($false); Release-Com $book; $book=$null
    $book=$books.Open($stageFile,0,$true)
    $reopened=Read-ReportState $book
    if (($updated | ConvertTo-Json -Depth 8 -Compress) -ne ($reopened | ConvertTo-Json -Depth 8 -Compress)) {
        throw 'Saved report counts or populated-cell/series counts did not survive reopening.'
    }
    $book.Close($false); Release-Com $book; $book=$null
    # File.Move refuses to overwrite even if another process creates Output now.
    [IO.File]::Move($stageFile,$outputPath)
    [ordered]@{output=$outputPath;updated=$updated;reopened=$reopened} | ConvertTo-Json -Depth 8
} finally {
    Release-Com $model
    if ($null -ne $book) { try { $book.Close($false) } catch {}; Release-Com $book }
    Release-Com $books
    if ($null -ne $excel) { try { $excel.Quit() } catch {}; Release-Com $excel }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    # Only the exact files this invocation owns; never recursively delete paths.
    if (Test-Path -LiteralPath $stageFile) { Remove-Item -LiteralPath $stageFile }
    if (Test-Path -LiteralPath $stageDirectory) { Remove-Item -LiteralPath $stageDirectory }
}
