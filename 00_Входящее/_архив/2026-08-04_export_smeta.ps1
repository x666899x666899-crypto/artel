param(
    [string]$src,
    [string]$out
)

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$wb = $null
try {
    $wb = $excel.Workbooks.Open($src, 0, $true)
    $i = 0
    foreach ($ws in $wb.Worksheets) {
        $i++
        $ur = $ws.UsedRange
        $rows = $ur.Rows.Count
        $cols = $ur.Columns.Count
        Write-Output ("SHEET {0}: {1} | rows={2} cols={3}" -f $i, $ws.Name, $rows, $cols)
        $vals = $ur.Value2
        $sb = New-Object System.Text.StringBuilder
        if ($vals -is [System.Array]) {
            for ($r = 1; $r -le $rows; $r++) {
                $line = New-Object System.Collections.Generic.List[string]
                for ($c = 1; $c -le $cols; $c++) {
                    $v = $vals.GetValue($r, $c)
                    if ($null -eq $v) { $line.Add("") } else { $line.Add([string]$v) }
                }
                [void]$sb.AppendLine(($line -join "`t"))
            }
        } else {
            [void]$sb.AppendLine([string]$vals)
        }
        $fname = Join-Path $out ("sheet" + $i + ".tsv")
        [System.IO.File]::WriteAllText($fname, $sb.ToString(), [System.Text.Encoding]::UTF8)
    }
} finally {
    if ($wb) { $wb.Close($false) }
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
}
Write-Output "DONE"
