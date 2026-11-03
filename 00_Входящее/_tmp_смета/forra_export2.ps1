# ASCII-only: parse xlsx (zip+xml) directly, no Excel COM.
Add-Type -AssemblyName System.IO.Compression.FileSystem

$desktop = "C:\Users\x6668\Desktop"
$artel = Get-ChildItem $desktop -Directory | ForEach-Object {
    Get-ChildItem $_.FullName -Directory -ErrorAction SilentlyContinue
} | Where-Object { $_.Name -eq "ARTEL" } | Select-Object -First 1
$tplDir = Get-ChildItem (Join-Path $artel.FullName "03_*") -Recurse -Directory |
    Where-Object { (Get-ChildItem $_.FullName -Filter *.xlsx -ErrorAction SilentlyContinue).Count -gt 30 } |
    Select-Object -First 1
$tmpParent = Get-ChildItem (Join-Path $artel.FullName "00_*") -Directory | Select-Object -First 1
$outDir = Join-Path $tmpParent.FullName "_forra_tsv"
New-Item -ItemType Directory -Force -Path $outDir | Out-Null

$books = @(Get-ChildItem $tplDir.FullName -Filter *.xlsx)
$confDir = Get-ChildItem $artel.FullName -Directory | Where-Object { $_.Name -like "99*" } | Select-Object -First 1
if ($confDir) { $books += @(Get-ChildItem $confDir.FullName -File | Where-Object { $_.Extension -eq ".xlsx" }) }
Write-Output ("BOOKS: " + $books.Count)

function ColIndex([string]$ref) {
    $letters = ($ref -replace "[0-9]", "")
    $n = 0
    foreach ($ch in $letters.ToCharArray()) { $n = $n * 26 + ([int]$ch - 64) }
    return $n
}

$done = 0; $failed = 0
foreach ($b in $books) {
    try {
        $zip = [System.IO.Compression.ZipFile]::OpenRead($b.FullName)
        try {
            # shared strings
            $shared = @()
            $e = $zip.GetEntry("xl/sharedStrings.xml")
            if ($e) {
                $sr = New-Object System.IO.StreamReader($e.Open())
                $xml = New-Object System.Xml.XmlDocument
                $xml.LoadXml($sr.ReadToEnd()); $sr.Close()
                $ns = New-Object System.Xml.XmlNamespaceManager($xml.NameTable)
                $ns.AddNamespace("m", $xml.DocumentElement.NamespaceURI)
                foreach ($si in $xml.SelectNodes("//m:si", $ns)) {
                    $txt = ""
                    foreach ($t in $si.SelectNodes(".//m:t", $ns)) { $txt += $t.InnerText }
                    $shared += $txt
                }
            }
            # sheet name -> file map
            $wbE = $zip.GetEntry("xl/workbook.xml")
            $sr = New-Object System.IO.StreamReader($wbE.Open())
            $wbXml = New-Object System.Xml.XmlDocument
            $wbXml.LoadXml($sr.ReadToEnd()); $sr.Close()
            $ns2 = New-Object System.Xml.XmlNamespaceManager($wbXml.NameTable)
            $ns2.AddNamespace("m", $wbXml.DocumentElement.NamespaceURI)
            $ns2.AddNamespace("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")
            $relE = $zip.GetEntry("xl/_rels/workbook.xml.rels")
            $sr = New-Object System.IO.StreamReader($relE.Open())
            $relXml = New-Object System.Xml.XmlDocument
            $relXml.LoadXml($sr.ReadToEnd()); $sr.Close()
            $relMap = @{}
            foreach ($rel in $relXml.DocumentElement.ChildNodes) { $relMap[$rel.GetAttribute("Id")] = $rel.GetAttribute("Target") }

            $sheetIdx = 0
            foreach ($sh in $wbXml.SelectNodes("//m:sheets/m:sheet", $ns2)) {
                $sheetIdx++
                $rid = $sh.GetAttribute("id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")
                $target = $relMap[$rid]
                if (-not $target) { continue }
                if ($target -notmatch "^xl/") { $target = "xl/" + $target }
                $se = $zip.GetEntry($target)
                if (-not $se) { continue }
                $sr = New-Object System.IO.StreamReader($se.Open())
                $sXml = New-Object System.Xml.XmlDocument
                $sXml.LoadXml($sr.ReadToEnd()); $sr.Close()
                $ns3 = New-Object System.Xml.XmlNamespaceManager($sXml.NameTable)
                $ns3.AddNamespace("m", $sXml.DocumentElement.NamespaceURI)
                $sb = New-Object System.Text.StringBuilder
                [void]$sb.AppendLine("### WORKBOOK: " + $b.BaseName + " | SHEET: " + $sh.GetAttribute("name"))
                foreach ($row in $sXml.SelectNodes("//m:sheetData/m:row", $ns3)) {
                    $rNum = $row.GetAttribute("r")
                    $cells = New-Object System.Collections.Generic.List[string]
                    $maxCol = 0
                    $vals = @{}
                    foreach ($c in $row.SelectNodes("m:c", $ns3)) {
                        $refAttr = $c.GetAttribute("r")
                        $ci = ColIndex $refAttr
                        if ($ci -gt $maxCol) { $maxCol = $ci }
                        $t = $c.GetAttribute("t")
                        $vNode = $c.SelectSingleNode("m:v", $ns3)
                        $v = ""
                        if ($t -eq "inlineStr") {
                            $isN = $c.SelectSingleNode("m:is", $ns3)
                            if ($isN) { $v = $isN.InnerText }
                        } elseif ($vNode) {
                            $v = $vNode.InnerText
                            if ($t -eq "s") { $v = $shared[[int]$v] }
                        }
                        if ($v -ne "") { $vals[$ci] = $v -replace "`t", " " -replace "`r?`n", " " }
                    }
                    if ($vals.Count -eq 0) { continue }
                    for ($i = 1; $i -le $maxCol; $i++) { if ($vals.ContainsKey($i)) { $cells.Add($vals[$i]) } else { $cells.Add("") } }
                    [void]$sb.AppendLine($rNum + "|" + ($cells -join "`t"))
                }
                $fname = Join-Path $outDir ($b.BaseName + "__" + $sheetIdx + ".tsv")
                [System.IO.File]::WriteAllText($fname, $sb.ToString(), [System.Text.Encoding]::UTF8)
            }
            $done++
        } finally { $zip.Dispose() }
    } catch {
        $failed++
        Write-Output ("FAIL: " + $b.Name + " | " + $_.Exception.Message)
    }
}
Write-Output ("DONE ok=" + $done + " fail=" + $failed)
