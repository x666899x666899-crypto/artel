# ASCII-only: aggregate unique positions from all Fora template TSVs.
$desktop = "C:\Users\x6668\Desktop"
$artel = Get-ChildItem $desktop -Directory | ForEach-Object {
    Get-ChildItem $_.FullName -Directory -ErrorAction SilentlyContinue
} | Where-Object { $_.Name -eq "ARTEL" } | Select-Object -First 1
$tmpParent = Get-ChildItem (Join-Path $artel.FullName "00_*") -Directory | Select-Object -First 1
$tsvDir = Join-Path $tmpParent.FullName "_forra_tsv"

$agg = @{}   # key -> object
$order = New-Object System.Collections.Generic.List[string]

foreach ($f in Get-ChildItem $tsvDir -Filter *.tsv) {
    $lines = [System.IO.File]::ReadAllLines($f.FullName, [System.Text.Encoding]::UTF8)
    if ($lines.Count -lt 2) { continue }
    $hdr = $lines[0]
    if ($hdr -notmatch "SHEET: (.+?)(\s*\|.*)?$") { continue }
    $sheet = $Matches[1].Trim()
    $wbName = ""
    if ($hdr -match "WORKBOOK: (.+?) \| SHEET") { $wbName = $Matches[1].Trim() }
    foreach ($line in $lines[1..($lines.Count - 1)]) {
        $p = $line.IndexOf("|")
        if ($p -lt 1) { continue }
        $cells = $line.Substring($p + 1) -split "`t"
        # data rows: [empty][code][name][...]; code is numeric
        if ($cells.Count -lt 4) { continue }
        if ($cells[1] -notmatch "^\d+$") { continue }
        $code = $cells[1]
        $name = $cells[2].Trim()
        if ($name -eq "") { continue }
        $rest = ($cells[3..([Math]::Min($cells.Count - 1, 5))] -join "`t")
        $key = $sheet + "|" + $code + "|" + $name
        if ($agg.ContainsKey($key)) {
            $agg[$key].Count++
        } else {
            $agg[$key] = [PSCustomObject]@{ Sheet = $sheet; Code = $code; Name = $name; Rest = $rest; Count = 1; First = $wbName }
            $order.Add($key)
        }
    }
}

$out = New-Object System.Text.StringBuilder
[void]$out.AppendLine("SHEET`tCODE`tNAME`tF1`tF2`tF3`tUSED_IN_N_TEMPLATES")
foreach ($k in $order) {
    $o = $agg[$k]
    [void]$out.AppendLine($o.Sheet + "`t" + $o.Code + "`t" + $o.Name + "`t" + $o.Rest + "`t" + $o.Count)
}
$outFile = Join-Path $tsvDir "_aggregate.tsv"
[System.IO.File]::WriteAllText($outFile, $out.ToString(), [System.Text.Encoding]::UTF8)
Write-Output ("ROWS: " + $order.Count + " -> " + $outFile)
