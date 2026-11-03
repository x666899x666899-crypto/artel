# ASCII-only script: exports every sheet of every Fora template xlsx to TSV
# Paths are discovered by wildcards to avoid Cyrillic literals.

$desktop = "C:\Users\x6668\Desktop"
$artel = Get-ChildItem $desktop -Directory | ForEach-Object {
    Get-ChildItem $_.FullName -Directory -ErrorAction SilentlyContinue
} | Where-Object { $_.Name -eq "ARTEL" } | Select-Object -First 1
if (-not $artel) { Write-Output "ARTEL root not found"; exit 1 }
Write-Output ("ARTEL: " + $artel.FullName)

# templates folder = dir under 03_* containing >30 xlsx files
$tplDir = Get-ChildItem (Join-Path $artel.FullName "03_*") -Recurse -Directory |
    Where-Object { (Get-ChildItem $_.FullName -Filter *.xlsx -ErrorAction SilentlyContinue).Count -gt 30 } |
    Select-Object -First 1
if (-not $tplDir) { Write-Output "Templates dir not found"; exit 1 }
Write-Output ("TPL: " + $tplDir.FullName)

# tmp output dir under 00_*
$tmpParent = Get-ChildItem (Join-Path $artel.FullName "00_*") -Directory | Select-Object -First 1
$outDir = Join-Path $tmpParent.FullName "_forra_tsv"
New-Item -ItemType Directory -Force -Path $outDir | Out-Null
Write-Output ("OUT: " + $outDir)

# collect workbooks: all templates + coefficient file from 99_*
$books = @(Get-ChildItem $tplDir.FullName -Filter *.xlsx)
$conf = Get-ChildItem (Join-Path $artel.FullName "99_*") -Filter *.xlsx -ErrorAction SilentlyContinue
$books += $conf
Write-Output ("BOOKS: " + $books.Count)

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
try { $excel.AskToUpdateLinks = $false } catch {}
$done = 0
$failed = 0

function Invoke-WithRetry([scriptblock]$action) {
    for ($try = 1; $try -le 6; $try++) {
        try { return (& $action) } catch {
            if ($_.Exception.Message -match "0x80010001" -and $try -lt 6) { Start-Sleep -Milliseconds (500 * $try) }
            else { throw }
        }
    }
}

try {
    foreach ($b in $books) {
        # skip books already exported on a previous run
        $existing = Get-ChildItem $outDir -Filter ($b.BaseName + "__*.tsv") -ErrorAction SilentlyContinue
        if ($existing) { $done++; continue }
        $wb = $null
        try {
            $wb = Invoke-WithRetry { $excel.Workbooks.Open($b.FullName, 0, $true) }
            $i = 0
            foreach ($ws in $wb.Worksheets) {
                $i++
                $ur = $ws.UsedRange
                $rows = $ur.Rows.Count
                $cols = $ur.Columns.Count
                if ($rows * $cols -gt 200000) { continue }
                $vals = Invoke-WithRetry { $ur.Value2 }
                $sb = New-Object System.Text.StringBuilder
                [void]$sb.AppendLine("### WORKBOOK: " + $b.BaseName + " | SHEET: " + $ws.Name + " | rows=" + $rows + " cols=" + $cols)
                if ($vals -is [System.Array]) {
                    for ($r = 1; $r -le $rows; $r++) {
                        $line = New-Object System.Collections.Generic.List[string]
                        $empty = $true
                        for ($c = 1; $c -le $cols; $c++) {
                            $v = $vals.GetValue($r, $c)
                            if ($null -eq $v) { $line.Add("") } else { $line.Add([string]$v); $empty = $false }
                        }
                        if (-not $empty) { [void]$sb.AppendLine($r.ToString() + "|" + ($line -join "`t")) }
                    }
                } else {
                    [void]$sb.AppendLine("1|" + [string]$vals)
                }
                $fname = Join-Path $outDir ($b.BaseName + "__" + $i + ".tsv")
                [System.IO.File]::WriteAllText($fname, $sb.ToString(), [System.Text.Encoding]::UTF8)
            }
            $done++
        } catch {
            $failed++
            Write-Output ("FAIL: " + $b.Name + " | " + $_.Exception.Message)
        } finally {
            if ($wb) { try { $wb.Close($false) } catch {} }
        }
    }
} finally {
    try { $excel.Quit() } catch {}
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
}
Write-Output ("DONE ok=" + $done + " fail=" + $failed)
