<# Native acceptance for synthetic analytical workbook fixtures. #>
[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)][string]$Source,
    [Parameter(Mandatory=$true)][string]$Output,
    [Parameter(Mandatory=$true)][string]$ExpectedCells
)
$ErrorActionPreference='Stop'
function Release-Com($value) {
    if ($null -ne $value -and [Runtime.InteropServices.Marshal]::IsComObject($value)) {
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($value)
    }
}
function Read-CheckedCells($book, $expectations) {
    $results=@()
    foreach ($expected in $expectations) {
        $sheet=$null; $cell=$null
        try {
            $sheet=$book.Worksheets.Item([string]$expected.sheet)
            $cell=$sheet.Range([string]$expected.cell)
            $actual=$cell.Value2
            $wanted=$expected.value
            if ($null -eq $wanted) { $matches=$null -eq $actual }
            elseif ($wanted -is [string]) { $matches=$actual -is [string] -and $actual -ceq $wanted }
            else { $matches=$null -ne $actual -and $actual -isnot [string] -and [Math]::Abs([double]$actual-[double]$wanted) -le 1e-9 }
            if (-not $matches) { throw "Cell $($expected.sheet)!$($expected.cell): expected '$wanted', observed '$actual' ($($cell.Text))" }
            $results+= [ordered]@{sheet=$expected.sheet;cell=$expected.cell;value=$actual;formula=$cell.Formula}
        } finally { Release-Com $cell; Release-Com $sheet }
    }
    return ,$results
}
$sourcePath=(Resolve-Path -LiteralPath $Source).Path
$outputPath=$ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Output)
if (Test-Path -LiteralPath $outputPath) { throw 'Output already exists.' }
if ([IO.Path]::GetExtension($sourcePath) -ne '.xlsx' -or [IO.Path]::GetExtension($outputPath) -ne '.xlsx') { throw 'Use .xlsx fixtures.' }
$spec=Get-Content -LiteralPath $ExpectedCells -Raw | ConvertFrom-Json
$expectations=@($spec)
if ($null -ne $spec.sheet_cells) {
    $expectations=@()
    foreach ($sheet in $spec.sheet_cells.PSObject.Properties) {
        foreach ($cell in $sheet.Value.PSObject.Properties) {
            $expectations+=@{sheet=$sheet.Name;cell=$cell.Name;value=$cell.Value}
        }
    }
}
if ($expectations.Count -eq 0) { throw 'At least one expected cell is required.' }
$excel=$null;$books=$null;$book=$null;$model=$null
try {
    Copy-Item -LiteralPath $sourcePath -Destination $outputPath
    $excel=New-Object -ComObject Excel.Application
    $excel.Visible=$false;$excel.DisplayAlerts=$false;$excel.AutomationSecurity=3;$excel.EnableEvents=$false
    $books=$excel.Workbooks
    $book=$books.Open($outputPath,0,$false)
    # CUBE formulas can initialize the embedded engine asynchronously during
    # Open. Excel briefly returns a workbook with no name/model interface.
    $ready=[Diagnostics.Stopwatch]::StartNew()
    while ([string]::IsNullOrEmpty($book.Name) -and $ready.Elapsed.TotalSeconds -lt 30) {
        Start-Sleep -Milliseconds 200
    }
    if ([string]::IsNullOrEmpty($book.Name)) { throw 'Excel did not finish opening the workbook within 30 seconds.' }
    $model=$book.Model;$model.Initialize();Release-Com $model;$model=$null
    $sheets=$null
    try {
        $sheets=$book.Worksheets
        for ($i=1;$i -le $sheets.Count;$i++) {
            $sheet=$null;$pivots=$null
            try {
                $sheet=$sheets.Item($i);$pivots=$sheet.PivotTables()
                for ($j=1;$j -le $pivots.Count;$j++) {
                    $pivot=$null
                    try { $pivot=$pivots.Item($j);$pivot.Update() }
                    finally { Release-Com $pivot }
                }
            } finally { Release-Com $pivots;Release-Com $sheet }
        }
    } finally { Release-Com $sheets }
    $excel.CalculateFullRebuild();$excel.CalculateUntilAsyncQueriesDone()
    $before=Read-CheckedCells $book $expectations
    $chartCount=0;$pivotTotals=@{}
    foreach ($sheet in $book.Worksheets) {
        $chartCount+=$sheet.ChartObjects().Count
        foreach ($pivot in $sheet.PivotTables()) {
            $values=@($pivot.TableRange1.Value2)
            $pivotTotals[$pivot.Name]=$values[-1]
        }
    }
    if ($null -ne $spec.chart_count -and $chartCount -ne $spec.chart_count) { throw 'Chart count mismatch' }
    if ($null -ne $spec.slicer_cache_count -and $book.SlicerCaches.Count -ne $spec.slicer_cache_count) { throw 'Control cache count mismatch' }
    if ($null -ne $spec.pivot_totals) {
        foreach ($pivot in $spec.pivot_totals.PSObject.Properties) {
            if ($pivotTotals[$pivot.Name] -ne $pivot.Value) { throw "Pivot total mismatch: $($pivot.Name)" }
        }
    }
    $book.Save();$book.Close($false);Release-Com $book;$book=$null
    $book=$books.Open($outputPath,0,$true)
    $ready.Restart()
    while ([string]::IsNullOrEmpty($book.Name) -and $ready.Elapsed.TotalSeconds -lt 30) {
        Start-Sleep -Milliseconds 200
    }
    if ([string]::IsNullOrEmpty($book.Name)) { throw 'Excel did not finish reopening the workbook within 30 seconds.' }
    $after=Read-CheckedCells $book $expectations
    [ordered]@{source=$sourcePath;output=$outputPath;passed=$true;before=$before;reopened=$after;chart_count=$chartCount;pivot_totals=$pivotTotals} |
        ConvertTo-Json -Depth 8 | Tee-Object -FilePath ($outputPath+'.json')
} finally {
    Release-Com $model
    if ($null -ne $book) { try { $book.Close($false) } catch {};Release-Com $book }
    Release-Com $books
    if ($null -ne $excel) { try { $excel.Quit() } catch {};Release-Com $excel }
    [GC]::Collect();[GC]::WaitForPendingFinalizers()
}
