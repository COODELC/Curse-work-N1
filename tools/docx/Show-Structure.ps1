param([string]$DocxPath = 'F:\курсовая Копейкин Андрей\Копейкин Андрей\Копейкин Андрей Последняя версия — копия\Копейкин _пояснительная_записка_курсового_проекта.docx')

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem

$zip = [System.IO.Compression.ZipFile]::OpenRead((Resolve-Path $DocxPath).Path)
try {
    $entry = $zip.Entries | Where-Object { $_.FullName -eq 'word/document.xml' }
    $reader = New-Object System.IO.StreamReader($entry.Open(), [System.Text.Encoding]::UTF8)
    $text = $reader.ReadToEnd()
    $reader.Close()
}
finally { $zip.Dispose() }

Write-Output '=== sectPr ==='
$m = [regex]::Match($text, '<w:sectPr[^>]*>.*?</w:sectPr>', 'Singleline')
Write-Output (($m.Value -replace '><', ">`n<"))

Write-Output ''
Write-Output '=== bookmarks with surrounding heading text ==='
foreach ($bm in [regex]::Matches($text, '<w:bookmarkStart w:id="(\d+)" w:name="(_Toc[^"]+)"')) {
    $id = $bm.Groups[1].Value
    $name = $bm.Groups[2].Value
    $after = $text.Substring($bm.Index, [Math]::Min(700, $text.Length - $bm.Index))
    $tm = [regex]::Match($after, '<w:t[^>]*>([^<]*)</w:t>')
    Write-Output ("id={0,-4} {1,-18} -> {2}" -f $id, $name, $tm.Groups[1].Value)
}

Write-Output ''
Write-Output '=== tblGrid samples ==='
$grids = [regex]::Matches($text, '<w:tblGrid>.*?</w:tblGrid>', 'Singleline')
$i = 0
foreach ($g in $grids) {
    $cols = ([regex]::Matches($g.Value, 'w:w="(\d+)"') | ForEach-Object { $_.Groups[1].Value }) -join '+'
    Write-Output ("grid[{0}] cols: {1} = {2}" -f $i, $cols, ([int](($cols -split '\+' | ForEach-Object { [int]$_ }) | Measure-Object -Sum).Sum))
    $i++
    if ($i -ge 12) { break }
}
