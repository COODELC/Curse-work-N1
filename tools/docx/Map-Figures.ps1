param([string]$ProjectRoot = 'F:\курсовая Копейкин Андрей\Копейкин Андрей\Копейкин Андрей Последняя версия — копия')
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem

$doc = Join-Path $ProjectRoot 'Копейкин _пояснительная_записка_курсового_проекта.docx'
$zip = [System.IO.Compression.ZipFile]::OpenRead($doc)
function Read-Entry($name) {
    $e = $zip.Entries | Where-Object { $_.FullName -eq $name }
    $sr = New-Object System.IO.StreamReader($e.Open(), [System.Text.Encoding]::UTF8)
    $s = $sr.ReadToEnd(); $sr.Close(); return $s
}
$docXml = Read-Entry 'word/document.xml'
$relsXml = Read-Entry 'word/_rels/document.xml.rels'
$zip.Dispose()

$relMap = @{}
foreach ($m in [regex]::Matches($relsXml, '<Relationship\s[^>]*Id="(rId\d+)"[^>]*Target="([^"]+)"')) {
    $relMap[$m.Groups[1].Value] = $m.Groups[2].Value
}
Write-Output ('Связей-изображений: ' + $relMap.Count)

$embeds = @()
foreach ($m in [regex]::Matches($docXml, '<a:blip[^>]*r:embed="(rId\d+)"')) {
    $embeds += $m.Groups[1].Value
}
Write-Output ('Встраиваний blip: ' + $embeds.Count)

$i = 0
foreach ($e in $embeds) {
    $i++
    $t = if ($relMap.ContainsKey($e)) { $relMap[$e] } else { '??? ' + $e }
    Write-Output ('  ris.{0,-3} -> {1}' -f $i, $t)
}
