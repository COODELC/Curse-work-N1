# Печатает XML элементов тела документа по индексу — для построения шаблонов правки.
#   powershell -File tools/docx/Dump-BodyXml.ps1 -DocxPath "..." -From 140 -To 165

param(
    [Parameter(Mandatory = $true)][string]$DocxPath,
    [int]$From = 0,
    [int]$To = 40,
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

$body = $document.SelectSingleNode('//w:body', $manager)

$lines = New-Object System.Collections.ArrayList
$index = 0

foreach ($node in $body.ChildNodes) {
    if ($index -ge $From -and $index -le $To) {
        $text = ($node.SelectNodes('.//w:t', $manager) | ForEach-Object { $_.InnerText }) -join ''

        [void]$lines.Add("========== [$index] $($node.LocalName) :: $text")
        [void]$lines.Add($node.OuterXml)
        [void]$lines.Add('')
    }

    $index++
}

$result = $lines -join "`r`n"

if ($OutFile -ne "") {
    $full = [System.IO.Path]::GetFullPath($OutFile)
    [System.IO.File]::WriteAllText($full, $result, (New-Object System.Text.UTF8Encoding($false)))
    Write-Output "XML сохранён: $full (элементов: $($To - $From + 1))"
} else {
    Write-Output $result
}
