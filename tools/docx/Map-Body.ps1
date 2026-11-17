# Карта элементов тела документа: индекс, тип, стиль, текст, признаки рисунка/таблицы.
#   powershell -File tools/docx/Map-Body.ps1 -DocxPath "..." -From 0 -To 400

param(
    [Parameter(Mandatory = $true)][string]$DocxPath,
    [int]$From = 0,
    [int]$To = 100000,
    [string]$OutFile = ""
)

$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName System.IO.Compression.FileSystem

$namespace = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

$zip = [System.IO.Compression.ZipFile]::OpenRead((Resolve-Path $DocxPath).Path)

try {
    $entry = $zip.Entries | Where-Object { $_.FullName -eq 'word/document.xml' }

    $reader = New-Object System.IO.StreamReader($entry.Open(), [System.Text.Encoding]::UTF8)
    $content = $reader.ReadToEnd()
    $reader.Close()
}
finally {
    $zip.Dispose()
}

$document = New-Object System.Xml.XmlDocument
$document.PreserveWhitespace = $true
$document.LoadXml($content)

$manager = New-Object System.Xml.XmlNamespaceManager($document.NameTable)
$manager.AddNamespace('w', $namespace)
$manager.AddNamespace('r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')

$lines = New-Object System.Collections.ArrayList
$index = 0

$body = $document.SelectSingleNode('//w:body', $manager)

foreach ($node in $body.ChildNodes) {
    if ($index -ge $From -and $index -le $To) {
        $style = $node.SelectSingleNode('w:pPr/w:pStyle/@w:val', $manager)
        $styleName = if ($style) { $style.Value } else { '-' }

        $text = ($node.SelectNodes('.//w:t', $manager) | ForEach-Object { $_.InnerText }) -join ''
        if ($text.Length -gt 70) { $text = $text.Substring(0, 70) + '...' }

        $flags = @()
        if ($node.SelectSingleNode('.//w:drawing', $manager)) { $flags += 'DRAWING' }
        if ($node.SelectSingleNode('.//w:pict', $manager)) { $flags += 'PICT' }
        if ($node.SelectSingleNode('.//w:br[@w:type="page"]', $manager)) { $flags += 'PAGEBREAK' }
        if ($node.LocalName -eq 'tbl') {
            $rows = $node.SelectNodes('w:tr', $manager).Count
            $styleEl = $node.SelectSingleNode('w:tblPr/w:tblStyle/@w:val', $manager)
            $flags += "ROWS=$rows"
            if ($styleEl) { $flags += "TBLSTYLE=$($styleEl.Value)" }
        }

        [void]$lines.Add(("[{0}] {1} <{2}> {3} | {4}" -f $index, $node.LocalName, $styleName, $text, ($flags -join ' ')))
    }

    $index++
}

$result = $lines -join "`r`n"

if ($OutFile -ne "") {
    $full = [System.IO.Path]::GetFullPath($OutFile)
    [System.IO.File]::WriteAllText($full, $result, (New-Object System.Text.UTF8Encoding($false)))
    Write-Output "Карта сохранена: $full (элементов: $($index))"
}
else {
    Write-Output $result
}
