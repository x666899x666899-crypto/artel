param([string]$src, [string]$out)
Add-Type -AssemblyName System.IO.Compression.FileSystem
$zip = [System.IO.Compression.ZipFile]::OpenRead($src)
try {
    $entry = $zip.Entries | Where-Object { $_.FullName -eq "word/document.xml" }
    $reader = New-Object System.IO.StreamReader($entry.Open(), [System.Text.Encoding]::UTF8)
    $xml = $reader.ReadToEnd()
    $reader.Close()
} finally {
    $zip.Dispose()
}
$xml = $xml -replace "<w:p [^>]*>", "`n" -replace "<w:p>", "`n" -replace "<w:tab[^>]*>", "`t" -replace "<w:br[^>]*>", "`n"
$text = $xml -replace "<[^>]+>", ""
$text = [System.Net.WebUtility]::HtmlDecode($text)
[System.IO.File]::WriteAllText($out, $text, [System.Text.Encoding]::UTF8)
Write-Output "EXTRACTED"
