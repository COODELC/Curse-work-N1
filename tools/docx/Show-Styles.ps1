param([string]$DocxPath = 'F:\курсовая Копейкин Андрей\Копейкин Андрей\Копейкин Андрей Последняя версия — копия\Копейкин _пояснительная_записка_курсового_проекта.docx')

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem

$zip = [System.IO.Compression.ZipFile]::OpenRead((Resolve-Path $DocxPath).Path)
try {
    $entry = $zip.Entries | Where-Object { $_.FullName -eq 'word/styles.xml' }
    $reader = New-Object System.IO.StreamReader($entry.Open(), [System.Text.Encoding]::UTF8)
    $text = $reader.ReadToEnd()
    $reader.Close()
}
finally { $zip.Dispose() }

Write-Output '=== docDefaults ==='
$i = $text.IndexOf('<w:docDefaults')
if ($i -ge 0) {
    $end = $text.IndexOf('</w:docDefaults>', $i)
    Write-Output (($text.Substring($i, $end - $i + 16) -replace '><', ">`n<"))
}

Write-Output ''
Write-Output '=== paragraph styles ==='
foreach ($m in [regex]::Matches($text, '<w:style w:type="paragraph"[^>]*w:styleId="([^"]+)"[^>]*>(.*?)</w:style>')) {
    $id = $m.Groups[1].Value
    $body = $m.Groups[2].Value
    $name = ''
    $nm = [regex]::Match($body, '<w:name w:val="([^"]+)"')
    if ($nm.Success) { $name = $nm.Groups[1].Value }

    $outline = ''
    $ol = [regex]::Match($body, '<w:outlineLvl w:val="(\d+)"')
    if ($ol.Success) { $outline = "outline=$($ol.Groups[1].Value)" }

    $based = ''
    $bs = [regex]::Match($body, '<w:basedOn w:val="([^"]+)"')
    if ($bs.Success) { $based = "basedOn=$($bs.Groups[1].Value)" }

    $num = ''
    $nm2 = [regex]::Match($body, '<w:numId w:val="(\d+)"')
    if ($nm2.Success) { $num = "numId=$($nm2.Groups[1].Value)" }

    $font = ''
    $fo = [regex]::Match($body, '<w:rFonts[^>]*w:ascii="([^"]+)"')
    if ($fo.Success) { $font = "font=$($fo.Groups[1].Value)" }

    $sz = ''
    $szm = [regex]::Match($body, '<w:sz w:val="(\d+)"')
    if ($szm.Success) { $sz = "sz=$($szm.Groups[1].Value)" }

    $jc = ''
    $jcm = [regex]::Match($body, '<w:jc w:val="([^"]+)"')
    if ($jcm.Success) { $jc = "jc=$($jcm.Groups[1].Value)" }

    Write-Output ("{0,-6} name='{1}' {2} {3} {4} {5} {6}" -f $id, $name, $based, $outline, $num, $font, "$sz $jc")
}

Write-Output ''
Write-Output '=== table styles ==='
foreach ($m in [regex]::Matches($text, '<w:style w:type="table"[^>]*w:styleId="([^"]+)"[^>]*>(.*?)</w:style>')) {
    Write-Output $m.Groups[1].Value
}
