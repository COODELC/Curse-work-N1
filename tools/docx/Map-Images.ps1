# Сопоставление рисунков: индекс элемента тела -> rId -> файл в word/media.
#   powershell -File tools/docx/Map-Images.ps1 -DocxPath "..." -OutFile "..."

param(
    [Parameter(Mandatory = $true)][string]$DocxPath,
    [string]$OutFile = ""
)

$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName System.IO.Compression.FileSystem

$nsW = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
$nsR = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
$nsPkg = 'http://schemas.openxmlformats.org/package/2006/relationships'

$zip = [System.IO.Compression.ZipFile]::OpenRead((Resolve-Path $DocxPath).Path)

function Read-Entry {
    param($Zip, [string]$Name)
    $entry = $Zip.Entries | Where-Object { $_.FullName -eq $Name }
    if (-not $entry) { return $null }
    $reader = New-Object System.IO.StreamReader($entry.Open(), [System.Text.Encoding]::UTF8)
    $text = $reader.ReadToEnd()
    $reader.Close()
    return $text
}

try {
    $docText = Read-Entry $zip 'word/document.xml'
    $relsText = Read-Entry $zip 'word/_rels/document.xml.rels'

    $relsDoc = New-Object System.Xml.XmlDocument
    $relsDoc.LoadXml($relsText)
    $relMap = @{}
    foreach ($rel in $relsDoc.DocumentElement.ChildNodes) {
        $relMap[$rel.GetAttribute('Id')] = $rel.GetAttribute('Target')
    }
}
finally {
    $zip.Dispose()
}

$document = New-Object System.Xml.XmlDocument
$document.PreserveWhitespace = $true
$document.LoadXml($docText)

$manager = New-Object System.Xml.XmlNamespaceManager($document.NameTable)
$manager.AddNamespace('w', $nsW)
$manager.AddNamespace('r', $nsR)

$mA = New-Object System.Xml.XmlNamespaceManager($document.NameTable)
$mA.AddNamespace('a', 'http://schemas.openxmlformats.org/drawingml/2006/main')
$mA.AddNamespace('r', $nsR)

$mW = New-Object System.Xml.XmlNamespaceManager($document.NameTable)
$mW.AddNamespace('wp', 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing')

$lines = New-Object System.Collections.ArrayList
$index = 0

$body = $document.SelectSingleNode('//w:body', $manager)

foreach ($node in $body.ChildNodes) {
    $blips = $node.SelectNodes('.//a:blip', $mA)

    if ($blips.Count -gt 0) {
        $ext = $node.SelectSingleNode('.//wp:inline/wp:extent', $mW)

        $size = if ($ext) { "$($ext.GetAttribute('cx'))x$($ext.GetAttribute('cy'))" } else { '' }

        $targets = @()
        foreach ($blip in $blips) {
            $rid = $blip.GetAttribute('embed', $nsR)
            $targets += "$rid=$($relMap[$rid])"
        }

        $text = ($node.SelectNodes('.//w:t', $manager) | ForEach-Object { $_.InnerText }) -join ''
        if ($text.Length -gt 50) { $text = $text.Substring(0, 50) + '...' }

        [void]$lines.Add(("[{0}] {1} | extent={2} | {3}" -f $index, ($targets -join ', '), $size, $text))
    }

    $index++
}

$result = $lines -join "`r`n"

if ($OutFile -ne "") {
    $full = [System.IO.Path]::GetFullPath($OutFile)
    [System.IO.File]::WriteAllText($full, $result, (New-Object System.Text.UTF8Encoding($false)))
    Write-Output "Сопоставление сохранено: $full"
}
else {
    Write-Output $result
}
