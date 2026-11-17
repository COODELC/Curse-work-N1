$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem
$root = 'F:\курсовая Копейкин Андрей\Копейкин Андрей\Копейкин Андрей Последняя версия — копия'
$work = Join-Path $root 'Копейкин _пояснительная_записка_курсового_проекта.docx'
$z = [System.IO.Compression.ZipFile]::OpenRead($work)
$e = $z.GetEntry('word/document.xml')
$ms = New-Object System.IO.MemoryStream
$s = $e.Open(); $s.CopyTo($ms); $s.Close(); $z.Dispose()
$docXml = [System.Text.Encoding]::UTF8.GetString($ms.ToArray())
$doc = New-Object System.Xml.XmlDocument
$doc.PreserveWhitespace = $true
$doc.LoadXml($docXml)
$ns = New-Object System.Xml.XmlNamespaceManager($doc.NameTable)
$ns.AddNamespace('w', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main')
$ns.AddNamespace('a', 'http://schemas.openxmlformats.org/drawingml/2006/main')
$ns.AddNamespace('r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
$body = $doc.SelectSingleNode('/w:document/w:body', $ns)
$paras = @($body.SelectNodes('.//w:p', $ns))

function Get-ParaText($p) {
    $parts = @($p.SelectNodes('.//w:t', $ns) | ForEach-Object { $_.InnerText })
    return ($parts -join '').Trim()
}

for ($i = 0; $i -lt $paras.Count; $i++) {
    $b = $paras[$i].SelectSingleNode('.//a:blip', $ns)
    if ($null -eq $b) { continue }
    $rid = $b.GetAttribute('embed', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
    $inTbl = ($null -ne $paras[$i].SelectSingleNode('ancestor::w:tbl', $ns))
    $lo = [Math]::Max(0, $i - 2)
    $hi = [Math]::Min($paras.Count - 1, $i + 3)
    Write-Output ("### $rid абз#$i вТаблице=$inTbl")
    for ($j = $lo; $j -le $hi; $j++) {
        $t = Get-ParaText $paras[$j]
        $bb = $paras[$j].SelectSingleNode('.//a:blip', $ns)
        $r = ''
        if ($null -ne $bb) { $r = ' <' + $bb.GetAttribute('embed', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships') + '>' }
        Write-Output ("   [$j]$r $t")
    }
}
