# Дамп структуры DOCX: абзацы со стилями и содержимое таблиц.
#   powershell -File tools/docx/Dump-Docx.ps1 -Input "doc.docx" -Output "docs\dump.txt"

param(
    [Parameter(Mandatory = $true)][string]$DocxPath,
    [string]$Output = "docs\docx_dump.txt"
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

function Get-ParagraphText {
    param([System.Xml.XmlNode]$Paragraph)

    $parts = New-Object System.Collections.ArrayList

    foreach ($node in $Paragraph.SelectNodes('.//w:t | .//w:tab | .//w:br', $manager)) {
        switch ($node.LocalName) {
            't' { [void]$parts.Add($node.InnerText) }
            'tab' { [void]$parts.Add(' ') }
            'br' { [void]$parts.Add(' / ') }
        }
    }

    return ($parts -join '')
}

$lines = New-Object System.Collections.ArrayList
$index = 0
$tableNumber = 0

$body = $document.SelectSingleNode('//w:body', $manager)

foreach ($node in $body.ChildNodes) {
    if ($node.LocalName -eq 'p') {
        $style = $node.SelectSingleNode('w:pPr/w:pStyle/@w:val', $manager)
        $styleName = if ($style) { $style.Value } else { '' }

        $outline = $node.SelectSingleNode('w:pPr/w:outlineLvl/@w:val', $manager)
        $outlineValue = if ($outline) { $outline.Value } else { '' }

        $text = Get-ParagraphText -Paragraph $node

        $numPr = $node.SelectSingleNode('w:pPr/w:numPr', $manager)
        $marker = if ($numPr) { '•' } else { '' }

        [void]$lines.Add(("[{0:D4}] <{1}> {2}{3}" -f $index, $styleName, $marker, $text))

        $index++
    }
    elseif ($node.LocalName -eq 'tbl') {
        $tableNumber++

        [void]$lines.Add('')
        [void]$lines.Add("=== ТАБЛИЦА $tableNumber ===")

        $rowNumber = 0

        foreach ($row in $node.SelectNodes('w:tr', $manager)) {
            $cells = New-Object System.Collections.ArrayList

            foreach ($cell in $row.SelectNodes('w:tc', $manager)) {
                $cellText = @()

                foreach ($paragraph in $cell.SelectNodes('w:p', $manager)) {
                    $cellText += (Get-ParagraphText -Paragraph $paragraph)
                }

                [void]$cells.Add(($cellText -join ' ⏎ '))
            }

            [void]$lines.Add(("  строка {0}: {1}" -f $rowNumber, ($cells -join ' || ')))

            $rowNumber++
        }

        [void]$lines.Add("=== КОНЕЦ ТАБЛИЦЫ $tableNumber ===")
        [void]$lines.Add('')
    }
}

$fullPath = [System.IO.Path]::GetFullPath($Output)
$dir = [System.IO.Path]::GetDirectoryName($fullPath)

if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Force -Path $dir | Out-Null }

[System.IO.File]::WriteAllText($fullPath, ($lines -join "`r`n"), (New-Object System.Text.UTF8Encoding($false)))

Write-Output "Абзацев: $index, таблиц: $tableNumber"
Write-Output "Дамп: $fullPath"
