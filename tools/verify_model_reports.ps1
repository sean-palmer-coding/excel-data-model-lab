param([string]$Source,[string]$Output,[string]$RefreshMode='Cache')
$ErrorActionPreference='Stop'
$excel=$null;$book=$null
function Report-State($b, $includeCharts=$true) {
    $result=@()
    foreach($ws in $b.Worksheets){
        foreach($p in $ws.PivotTables()){
            $orients=@();foreach($cf in $p.CubeFields){if($cf.Orientation -ne 0){$orients+=@{name=$cf.Name;axis=$cf.Orientation}};[void][Runtime.InteropServices.Marshal]::ReleaseComObject($cf)}
            $result+=@{sheet=$ws.Name;name=$p.Name;olap=$p.PivotCache().OLAP;range=$p.TableRange1.Address();values=@($p.TableRange1.Value2);axes=$orients;manual=$p.ManualUpdate}
            [void][Runtime.InteropServices.Marshal]::ReleaseComObject($p)
        }
        if($includeCharts){foreach($co in $ws.ChartObjects()){
            $cs=$co.Chart.SeriesCollection();$count=0;if($null -ne $cs){$count=$cs.Count}
            $result+=@{chart=$co.Name;sheet=$ws.Name;pivot=$co.Chart.PivotLayout.PivotTable.Name;series=$count;type=$co.Chart.ChartType}
            [void][Runtime.InteropServices.Marshal]::ReleaseComObject($co)
        }}
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($ws)
    }
    return $result
}
try{
    if([IO.Path]::GetExtension($Source) -ne '.xlsx' -or [IO.Path]::GetExtension($Output) -ne '.xlsx'){throw 'Source and Output must be .xlsx workbooks'}
    if(Test-Path -LiteralPath $Output){throw 'Native verification output exists'}
    Copy-Item -LiteralPath $Source -Destination $Output
    $excel=New-Object -ComObject Excel.Application
    # COM-started Excel enables macros by default; force-disable them and events.
    $excel.Visible=$false;$excel.DisplayAlerts=$false;$excel.AutomationSecurity=3;$excel.EnableEvents=$false
    $book=$excel.Workbooks.Open($Output,0,$false)
    $book.Model.Initialize()
    $result=@{source=$Source;opened=Report-State $book $false}
    if($RefreshMode -eq 'All'){
        $book.RefreshAll()
        $excel.CalculateUntilAsyncQueriesDone()
    } else { foreach($ws in $book.Worksheets){
        foreach($p in $ws.PivotTables()){
            if($RefreshMode -eq 'CacheOnly'){$p.PivotCache().Refresh()}
            elseif($RefreshMode -eq 'UpdateOnly'){$p.Update()}
            else {
                if(-not $p.RefreshTable()){throw 'Pivot refresh returned false'}
                $p.PivotCache().Refresh()
                $p.Update()
            }
            [void][Runtime.InteropServices.Marshal]::ReleaseComObject($p)
        }
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($ws)
    } }
    $result.refreshed=Report-State $book
    $result | ConvertTo-Json -Depth 9 | Set-Content -LiteralPath ($Output+'.partial.json')
    $book.Save();$book.Close($false)
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($book);$book=$null
    [GC]::Collect();[GC]::WaitForPendingFinalizers()
    $book=$excel.Workbooks.Open($Output,0,$true)
    $result.reopened=Report-State $book
    $result | ConvertTo-Json -Depth 9 | Set-Content -LiteralPath ($Output+'.json')
    $result | ConvertTo-Json -Depth 9 -Compress | Write-Output
} finally {
    if($null -ne $book){try{$book.Close($false)}catch{};[void][Runtime.InteropServices.Marshal]::ReleaseComObject($book)}
    if($null -ne $excel){try{$excel.Quit()}catch{};[void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel)}
    [GC]::Collect();[GC]::WaitForPendingFinalizers()
}
