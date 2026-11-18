<#
Verify real Excel controls, filtering, clearing and save/reopen in one owned
Excel instance. Source is copied; an existing output is never overwritten.
For examples/model_reports.py use -Example. Other fixtures can supply expected
total maps (PivotTable name => bottom-right grand total). These assertions
require a single measure with row/column grand totals enabled.
#>
param(
    [Parameter(Mandatory=$true)][string]$Source,
    [Parameter(Mandatory=$true)][string]$Output,
    [string]$CategoryTable='Products',
    [string]$CategoryColumn='Category',
    [string]$CategoryMember='Alpha',
    [datetime]$StartDate=[datetime]'2026-01-01',
    [datetime]$EndDate=[datetime]'2026-01-31',
    [hashtable]$ExpectedInitial=@{},
    [hashtable]$ExpectedSlicer=@{},
    [hashtable]$ExpectedTimeline=@{},
    [int]$ExpectedControlCount=2,
    [int]$ExpectedConnections=2,
    [switch]$Example
)
$ErrorActionPreference='Stop'
$excel=$null;$book=$null
if($Example){
    $CategoryTable='Sales'
    $ExpectedInitial=@{SalesPivot=100;TotalPivot=100;EastPivot=30}
    $ExpectedSlicer=@{SalesPivot=40;TotalPivot=40;EastPivot=30}
    $ExpectedTimeline=@{SalesPivot=30;TotalPivot=30;EastPivot=30}
}
function Snapshot($b){
    $rows=@()
    foreach($ws in $b.Worksheets){
        foreach($p in $ws.PivotTables()){
            $range=$p.TableRange1
            $last=$range.Cells.Item($range.Rows.Count,$range.Columns.Count)
            $rows+=@{name=$p.Name;values=@($range.Value2);total=$last.Value2}
            foreach($o in @($last,$range,$p)){[void][Runtime.InteropServices.Marshal]::ReleaseComObject($o)}
        }
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($ws)
    }
    return $rows
}
function Assert-Totals($rows,[hashtable]$expected,[string]$stage){
    foreach($name in $expected.Keys){
        $found=@($rows|Where-Object {$_.name -eq $name})
        if($found.Count -ne 1){throw "${stage}: expected exactly one PivotTable named $name"}
        if($null -eq $found[0].total -or [double]$found[0].total -ne [double]$expected[$name]){
            throw "${stage}: $name total '$($found[0].total)' differs from '$($expected[$name])'"
        }
    }
}
function Escape-Member([string]$value){return $value.Replace(']',']]')}
try {
    $Source=(Resolve-Path -LiteralPath $Source).Path
    $Output=[IO.Path]::GetFullPath($Output)
    if([IO.Path]::GetExtension($Source) -ne '.xlsx' -or [IO.Path]::GetExtension($Output) -ne '.xlsx'){throw 'Source and Output must be .xlsx workbooks'}
    if(Test-Path -LiteralPath $Output){throw 'Verification output exists'}
    Copy-Item -LiteralPath $Source -Destination $Output
    $excel=New-Object -ComObject Excel.Application
    # COM-started Excel enables macros by default; force-disable them and events.
    $excel.Visible=$false;$excel.DisplayAlerts=$false;$excel.AutomationSecurity=3;$excel.EnableEvents=$false
    $book=$excel.Workbooks.Open($Output,0,$false)
    $book.Model.Initialize()
    foreach($ws in $book.Worksheets){
        foreach($p in $ws.PivotTables()){
            $p.RefreshTable()|Out-Null;$p.PivotCache().Refresh();$p.Update()
            [void][Runtime.InteropServices.Marshal]::ReleaseComObject($p)
        }
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($ws)
    }
    $result=@{source=$Source;initial=Snapshot $book;controls=@()}
    Assert-Totals $result.initial $ExpectedInitial 'initial'
    if($book.SlicerCaches.Count -ne $ExpectedControlCount){throw 'Unexpected control cache count'}
    $types=@()
    foreach($sc in $book.SlicerCaches){
        $info=@{name=$sc.Name;source=$sc.SourceName;type=$sc.SlicerCacheType;views=$sc.Slicers.Count;connections=$sc.PivotTables.Count}
        if($info.views -lt 1 -or $info.connections -ne $ExpectedConnections){throw "Missing view or incorrect connections for $($info.name)"}
        $types+=$info.type
        if($info.type -eq 1){
            $sc.VisibleSlicerItemsList=@("[$(Escape-Member $CategoryTable)].[$(Escape-Member $CategoryColumn)].&[$(Escape-Member $CategoryMember)]")
            $info.filtered=Snapshot $book
            Assert-Totals $info.filtered $ExpectedSlicer 'slicer'
            $sc.ClearManualFilter()
        }elseif($info.type -eq 2){
            $sc.TimelineState.SetFilterDateRange($StartDate,$EndDate)
            $info.filtered=Snapshot $book
            Assert-Totals $info.filtered $ExpectedTimeline 'timeline'
            $sc.ClearDateFilter()
        }else{throw "Unknown control type $($info.type)"}
        Assert-Totals (Snapshot $book) $ExpectedInitial 'cleared control'
        $result.controls+=$info
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($sc)
    }
    if($Example -and (($types -notcontains 1) -or ($types -notcontains 2))){throw 'Example requires both a slicer and timeline'}
    $result.cleared=Snapshot $book
    $book.Save();$book.Close($false)
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($book);$book=$null
    $book=$excel.Workbooks.Open($Output,0,$true)
    $result.reopened=Snapshot $book
    Assert-Totals $result.reopened $ExpectedInitial 'reopened'
    if($book.SlicerCaches.Count -ne $ExpectedControlCount){throw 'Controls missing after reopen'}
    $result.passed=$true
    $result|ConvertTo-Json -Depth 9|Set-Content -LiteralPath ($Output+'.json')
    $result|ConvertTo-Json -Depth 9 -Compress|Write-Output
} finally {
    if($null -ne $book){try{$book.Close($false)}catch{};[void][Runtime.InteropServices.Marshal]::ReleaseComObject($book)}
    if($null -ne $excel){try{$excel.Quit()}catch{};[void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel)}
    [GC]::Collect();[GC]::WaitForPendingFinalizers()
}
