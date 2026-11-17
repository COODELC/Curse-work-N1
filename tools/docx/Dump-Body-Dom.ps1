param(
    [Parameter(Mandatory = $true)][string]$Docx,
    [int]$From = 0,
    [int]$To = 100000
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
$i = -1
foreach ($n in $body.ChildNodes) {
    if ($n.NodeType -ne 'Element') { continue }
    if ($n.LocalName -ne 'p' -and $n.LocalName -ne 'tbl') { continue }
    $i++
    if ($i -lt $From -or $i -gt $To) { continue }
    $flags = ''
    if ($n.SelectNodes('.//w:drawing', $ns).Count -gt 0) { $flags += ' DRAWING' }
    if ($n.SelectNodes('.//w:object', $ns).Count -gt 0) { $flags += ' OBJECT' }
    if ($n.LocalName -eq 'tbl') {
        $rows = $n.SelectNodes('./w:tr', $ns).Count
        Write-Output ("[$i] TABLE rows=$rows$flags")
        continue
    }
    $sb = New-Object System.Text.StringBuilder
    foreach ($t in $n.SelectNodes('.//w:t', $ns)) { [void]$sb.Append($t.InnerText) }
    Write-Output ("[$i]$flags`t" + $sb.ToString())
}
