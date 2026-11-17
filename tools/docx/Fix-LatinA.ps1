$ErrorActionPreference = 'Stop'
$root = 'F:\курсовая Копейкин Андрей\Копейкин Андрей\Копейкин Андрей Последняя версия — копия'
$docPath = Join-Path $root 'Копейкин _пояснительная_записка_курсового_проекта.docx'
$stamp = Get-Date -Format 'yyyy-MM-dd-HHmmss'
$bak = Join-Path $root ('docs\backup_zapiska_' + $stamp + '.docx')
Copy-Item -LiteralPath $docPath -Destination $bak -Force
Write-Output ('Бэкап: ' + $bak)

Add-Type -AssemblyName System.IO.Compression.FileSystem
$W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
$zip = [System.IO.Compression.ZipFile]::Open($docPath, 'Update')
try {
    $entry = $zip.Entries | Where-Object { $_.FullName -eq 'word/document.xml' }
    $stream = $entry.Open()
    $sr = New-Object System.IO.StreamReader($stream, [System.Text.Encoding]::UTF8)
    $content = $sr.ReadToEnd()
    $sr.Close()

    $xml = New-Object System.Xml.XmlDocument
    $xml.PreserveWhitespace = $true
    $xml.LoadXml($content)
    $ns = New-Object System.Xml.XmlNamespaceManager($xml.NameTable)
    $ns.AddNamespace('w', $W)
    $body = $xml.SelectSingleNode('/w:document/w:body', $ns)

    function Get-PText($p) {
        $sb = New-Object System.Text.StringBuilder
        foreach ($t in $p.SelectNodes('.//w:t', $ns)) { [void]$sb.Append($t.InnerText) }
        return $sb.ToString()
    }

    $cyr = [string][char]1040
    $done = 0
    foreach ($p in @($body.SelectNodes('.//w:p', $ns))) {
        $pt = Get-PText $p
        if ($pt -notmatch 'Рисунок\s+A\.1') { continue }
        $replaced = $false
        foreach ($t in @($p.SelectNodes('.//w:t', $ns))) {
            if ($replaced) { break }
            $v = $t.InnerText
            $idx = $v.IndexOf('A')
            if ($idx -ge 0) {
                $t.InnerText = $v.Substring(0, $idx) + $cyr + $v.Substring($idx + 1)
                $replaced = $true
            }
        }
        if ($replaced) { $done++; Write-Output ('  исправлено: [' + (Get-PText $p).Trim() + ']') }
    }
    Write-Output ('Исправлено подписей: ' + $done)

    # проверка: не осталось ли латинской A в обозначениях приложений
    $rest = 0
    foreach ($p in @($body.SelectNodes('.//w:p', $ns))) {
        $pt = Get-PText $p
        if ($pt -match 'ПРИЛОЖЕНИЕ\s+A' -or $pt -match 'Рисунок\s+A\.1') { $rest++ }
    }
    Write-Output ('Осталось латинских A: ' + $rest)

    $settings = New-Object System.Xml.XmlWriterSettings
    $settings.Encoding = New-Object System.Text.UTF8Encoding($false)
    $settings.Indent = $false
    $s2 = $entry.Open()
    $s2.SetLength(0)
    $writer = [System.Xml.XmlWriter]::Create($s2, $settings)
    $xml.Save($writer)
    $writer.Flush(); $writer.Close(); $s2.Close()
    Write-Output 'XML записан'
}
finally { $zip.Dispose() }
Write-Output 'OK'
