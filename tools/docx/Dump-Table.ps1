param(
    [Parameter(Mandatory = $true)][string]$Docx,
    [string]$Caption = ''
)

Add-Type -AssemblyName System.IO.Compression.FileSystem
$ErrorActionPreference = 'Stop'

$zip = [System.IO.Compression.ZipFile]::OpenRead($Docx)
try {
    $entry = $zip.GetEntry('word/document.xml')
    $sr = New-Object System.IO.StreamReader($entry.Open(), [Text.Encoding]::UTF8)
    $xml = $sr.ReadToEnd()
    $sr.Close()
}
finally { $zip.Dispose() }

$doc = New-Object System.Xml.XmlDocument
$doc.PreserveWhitespace = $true
$doc.LoadXml($xml)
$ns = New-Object System.Xml.XmlNamespaceManager($doc.NameTable)
$ns.AddNamespace('w', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main')

$body = $doc.SelectSingleNode('//w:body', $ns)
$captionHit = ($Caption -eq '')

foreach ($n in $body.ChildNodes) {
    if ($n.NodeType -ne 'Element') { continue }
    if ($n.LocalName -eq 'p' -and $Caption -ne '') {
        $t = ''
        foreach ($tt in $n.SelectNodes('.//w:t', $ns)) { $t += $tt.InnerText }
        if ($t.Trim().Contains($Caption)) { $captionHit = $true }
        continue
    }
    if ($n.LocalName -ne 'tbl' -or -not $captionHit) { continue }
    $captionHit = $false
    Write-Output '--- TABLE ---'
    foreach ($tr in $n.SelectNodes('./w:tr', $ns)) {
        $cells = @()
        foreach ($tc in $tr.SelectNodes('./w:tc', $ns)) {
            $t = ''
            foreach ($tt in $tc.SelectNodes('.//w:t', $ns)) { $t += $tt.InnerText }
            $cells += $t.Trim()
        }
        Write-Output ('| ' + ($cells -join ' | '))
    }
}
