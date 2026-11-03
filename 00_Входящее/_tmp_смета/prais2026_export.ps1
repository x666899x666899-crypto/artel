# ASCII-only: export price xlsx -> TSV + markdown (no Cyrillic in source)
Add-Type -AssemblyName System.IO.Compression.FileSystem

$desktop = "C:\Users\x6668\Desktop"
$artel = Get-ChildItem $desktop -Directory | ForEach-Object {
    Get-ChildItem $_.FullName -Directory -ErrorAction SilentlyContinue
} | Where-Object { $_.Name -eq "ARTEL" } | Select-Object -First 1
if (-not $artel) { Write-Output "NO_ARTEL"; exit 1 }

$incoming = Get-ChildItem $artel.FullName -Directory | Where-Object { $_.Name -like "00*" } | Select-Object -First 1
$fin = Get-ChildItem $artel.FullName -Directory | Where-Object { $_.Name -like "10*" } | Select-Object -First 1
$outMd = Get-ChildItem $fin.FullName -Directory | Where-Object { $_.Name -like "*райс*" } | Select-Object -First 1
if (-not $outMd) { Write-Output "NO_PRICES_DIR"; exit 1 }
$outTsv = Join-Path $incoming.FullName "_prais2026_tsv"
New-Item -ItemType Directory -Force -Path $outTsv | Out-Null

$xlsx = Get-ChildItem -LiteralPath $incoming.FullName -Filter *.xlsx -ErrorAction SilentlyContinue |
    Sort-Object LastWriteTime -Descending | Select-Object -First 1
if (-not $xlsx) { Write-Output "NO_FILE"; exit 1 }
Write-Output ("SOURCE: " + $xlsx.FullName)

function ColIndex([string]$ref) {
    $letters = ($ref -replace "[0-9]", "")
    $n = 0
    foreach ($ch in $letters.ToCharArray()) { $n = $n * 26 + ([int]$ch - 64) }
    return $n
}

function Slug([string]$name) {
    $s = ($name.ToLower() -replace "[^a-z0-9\u0400-\u04ff]+", "_").Trim("_")
    if ($s.Length -gt 40) { $s = $s.Substring(0, 40) }
    return $s
}

function TsvToMarkdown([string[]]$lines) {
    $rows = @()
    foreach ($line in $lines) { if ($line -match "^\d+\|") { $rows += ($line -replace "^\d+\|", "") } }
    if ($rows.Count -eq 0) { return @("", 0) }
    $parsed = @(); foreach ($r in $rows) { $parsed += ,@($r -split "`t", -1) }
    $maxCol = ($parsed | ForEach-Object { $_.Count } | Measure-Object -Maximum).Maximum
    while ($maxCol -gt 1) {
        $empty = $true
        foreach ($p in $parsed) { if ($p.Count -ge $maxCol -and $p[$maxCol-1]) { $empty = $false; break } }
        if ($empty) { $maxCol-- } else { break }
    }
    $norm = @()
    foreach ($p in $parsed) {
        $c = @($p); while ($c.Count -lt $maxCol) { $c += "" }
        if ($c.Count -gt $maxCol) { $c = $c[0..($maxCol-1)] }
        if (($c -join "").Trim()) { $norm += ,@($c) }
    }
    if ($norm.Count -eq 0) { return @("", 0) }
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine("| " + ($norm[0] -join " | ") + " |")
    [void]$sb.AppendLine("| " + ((1..$maxCol | ForEach-Object { "---" }) -join " | ") + " |")
    for ($i = 1; $i -lt $norm.Count; $i++) { [void]$sb.AppendLine("| " + ($norm[$i] -join " | ") + " |") }
    return @($sb.ToString(), $norm.Count)
}

$zip = [System.IO.Compression.ZipFile]::OpenRead($xlsx.FullName)
$sheetMeta = @()
try {
    $shared = @()
    $e = $zip.GetEntry("xl/sharedStrings.xml")
    if ($e) {
        $sr = New-Object System.IO.StreamReader($e.Open())
        $xml = New-Object System.Xml.XmlDocument; $xml.LoadXml($sr.ReadToEnd()); $sr.Close()
        $ns = New-Object System.Xml.XmlNamespaceManager($xml.NameTable)
        $ns.AddNamespace("m", $xml.DocumentElement.NamespaceURI)
        foreach ($si in $xml.SelectNodes("//m:si", $ns)) {
            $txt = ""; foreach ($t in $si.SelectNodes(".//m:t", $ns)) { $txt += $t.InnerText }
            $shared += $txt
        }
    }
    $wbE = $zip.GetEntry("xl/workbook.xml")
    $sr = New-Object System.IO.StreamReader($wbE.Open())
    $wbXml = New-Object System.Xml.XmlDocument; $wbXml.LoadXml($sr.ReadToEnd()); $sr.Close()
    $ns2 = New-Object System.Xml.XmlNamespaceManager($wbXml.NameTable)
    $ns2.AddNamespace("m", $wbXml.DocumentElement.NamespaceURI)
    $relE = $zip.GetEntry("xl/_rels/workbook.xml.rels")
    $sr = New-Object System.IO.StreamReader($relE.Open())
    $relXml = New-Object System.Xml.XmlDocument; $relXml.LoadXml($sr.ReadToEnd()); $sr.Close()
    $relMap = @{}; foreach ($rel in $relXml.DocumentElement.ChildNodes) { $relMap[$rel.GetAttribute("Id")] = $rel.GetAttribute("Target") }

    $sheetIdx = 0
    foreach ($sh in $wbXml.SelectNodes("//m:sheets/m:sheet", $ns2)) {
        $sheetIdx++
        $sName = $sh.GetAttribute("name")
        $rid = $sh.GetAttribute("id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")
        $target = $relMap[$rid]; if ($target -notmatch "^xl/") { $target = "xl/" + $target }
        $se = $zip.GetEntry($target); if (-not $se) { continue }
        $sr = New-Object System.IO.StreamReader($se.Open())
        $sXml = New-Object System.Xml.XmlDocument; $sXml.LoadXml($sr.ReadToEnd()); $sr.Close()
        $ns3 = New-Object System.Xml.XmlNamespaceManager($sXml.NameTable)
        $ns3.AddNamespace("m", $sXml.DocumentElement.NamespaceURI)
        $lines = New-Object System.Collections.Generic.List[string]
        [void]$lines.Add("SHEET:" + $sName)
        foreach ($row in $sXml.SelectNodes("//m:sheetData/m:row", $ns3)) {
            $rNum = $row.GetAttribute("r"); $maxCol = 0; $vals = @{}
            foreach ($c in $row.SelectNodes("m:c", $ns3)) {
                $ci = ColIndex ($c.GetAttribute("r")); if ($ci -gt $maxCol) { $maxCol = $ci }
                $t = $c.GetAttribute("t"); $vNode = $c.SelectSingleNode("m:v", $ns3); $v = ""
                if ($t -eq "inlineStr") { $isN = $c.SelectSingleNode("m:is", $ns3); if ($isN) { $v = $isN.InnerText } }
                elseif ($vNode) { $v = $vNode.InnerText; if ($t -eq "s") { $v = $shared[[int]$v] } }
                if ($v) { $vals[$ci] = $v -replace "`t", " " -replace "`r?`n", " " }
            }
            if ($vals.Count -eq 0) { continue }
            $cells = New-Object System.Collections.Generic.List[string]
            for ($i = 1; $i -le $maxCol; $i++) { [void]$cells.Add($(if ($vals.ContainsKey($i)) { $vals[$i] } else { "" })) }
            [void]$lines.Add($rNum + "|" + ($cells -join "`t"))
        }
        $tsvPath = Join-Path $outTsv ($sheetIdx.ToString("00") + "_" + (Slug $sName) + ".tsv")
        [System.IO.File]::WriteAllText($tsvPath, ($lines -join "`n"), [System.Text.Encoding]::UTF8)

        $skip = ($sName -match "coglasovan" -or $sName -match "\u0441\u043e\u0433\u043b\u0430\u0441\u043e\u0432\u0430\u043d")
        $mdBody = ""; $rowCount = 0
        if (-not $skip) { $res = TsvToMarkdown $lines; $mdBody = $res[0]; $rowCount = $res[1] }
        $mdName = "Price_2026_" + $sheetIdx.ToString("00") + "_" + (Slug $sName) + ".md"
        $mdPath = Join-Path $outMd.FullName $mdName
        $hdr = "# Price 2026 / " + $sName + "`n`n> Source: Google Sheet 1giAkQEGQXNBjSbnmejQIRmUgl7rpPj4rpwCgDSyGJd0`n> Sheet: " + $sName + "`n> File: " + $xlsx.Name + "`n> Date: " + (Get-Date -Format "yyyy-MM-dd") + "`n`n"
        [System.IO.File]::WriteAllText($mdPath, $hdr + $(if ($mdBody) { $mdBody } else { "_service sheet_`n" }), [System.Text.Encoding]::UTF8)
        $sheetMeta += [PSCustomObject]@{ Sheet = $sName; Rows = $rowCount; File = $mdName }
        Write-Output ("OK " + $sheetIdx + " " + $sName)
    }
} finally { $zip.Dispose() }

$index = Join-Path $outMd.FullName "Price_2026_INDEX.md"
$idx = @("# Price 2026 index", "", "| Sheet | File | Rows |", "|------|------|------|")
foreach ($m in $sheetMeta) { $idx += "| " + $m.Sheet + " | [" + $m.File + "](./" + $m.File + ") | " + $m.Rows + " |" }
[System.IO.File]::WriteAllText($index, ($idx -join "`n"), [System.Text.Encoding]::UTF8)
Write-Output "DONE"
