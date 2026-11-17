param(
    [string]$DocxPath = 'F:\курсовая Копейкин Андрей\Копейкин Андрей\Копейкин Андрей Последняя версия — копия\Копейкин _пояснительная_записка_курсового_проекта.docx',
    [int[]]$Indexes = @()
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem

$nsW = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

$zip = [System.IO.Compression.ZipFile]::OpenRead((Resolve-Path $DocxPath).Path)
try {
    $entry = $zip.Entries | Where-Object { $_.FullName -eq 'word/document.xml' }
    $reader = New-Object System.IO.StreamReader($entry.Open(), [System.Text.Encoding]::UTF8)
    $text = $reader.ReadToEnd()
    $reader.Close()
}
finally { $zip.Dispose() }

$document = New-Object System.Xml.XmlDocument
$document.PreserveWhitespace = $true
$document.LoadXml($text)

$manager = New-Object System.Xml.XmlNamespaceManager($document.NameTable)
$manager.AddNamespace('w', $nsW)

$body = $document.SelectSingleNode('//w:body', $manager)

$index = 0

foreach ($node in $body.ChildNodes) {
    if ($Indexes -contains $index) {
        Write-Output "========== [$index] $($node.LocalName) =========="
        Write-Output $node.OuterXml
        Write-Output ''
    }

    $index++
}

Write-Output '########## абзацы с полем SEQ ##########'

$index = 0

foreach ($node in $body.ChildNodes) {
    if ($node.LocalName -ne 'p') { $index++; continue }

    $instr = $node.SelectNodes('.//w:instrText', $manager)

    foreach ($i in $instr) {
        if ($i.InnerText -like '*SEQ*') {
            $txt = ($node.SelectNodes('.//w:t', $manager) | ForEach-Object { $_.InnerText }) -join ''
            Write-Output "[$index] SEQ: $($i.InnerText.Trim()) | текст: $txt"
            Write-Output $node.OuterXml
            Write-Output ''
        }
    }

    $index++
}
