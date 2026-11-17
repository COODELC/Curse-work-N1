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

    # ---------- 1. Удаление дублирующихся подписей «Таблица N» ----------
    $children = @($body.ChildNodes)
    $paras = @($body.SelectNodes('w:p', $ns))
    $toRemove = New-Object System.Collections.ArrayList

    for ($i = 0; $i -lt $paras.Count; $i++) {
        $t = (Get-PText $paras[$i]).Trim()
        if ($t -match '^Таблица\s+(\d+)$') {
            $num = $matches[1]
            # ищем соседа по элементам body
            $idx = [array]::IndexOf($children, $paras[$i])
            $prevT = ''
            $nextT = ''
            for ($j = $idx - 1; $j -ge 0; $j--) {
                if ($children[$j].LocalName -eq 'p') {
                    $prevT = (Get-PText $children[$j]).Trim()
                    if ($prevT -ne '') { break }
                }
            }
            for ($j = $idx + 1; $j -lt $children.Count; $j++) {
                if ($children[$j].LocalName -eq 'p') {
                    $nextT = (Get-PText $children[$j]).Trim()
                    if ($nextT -ne '') { break }
                }
            }
            $pat = '^Таблица\s+' + $num + '\s+[\u2014\u2013-]'
            if (($nextT -match $pat) -or ($prevT -match $pat)) {
                [void]$toRemove.Add($paras[$i])
                Write-Output ('  дубль: [' + $t + '] сосед: [' + $(if ($nextT -match $pat) { $nextT } else { $prevT }) + ']')
            }
        }
    }
    Write-Output ('К удалению: ' + $toRemove.Count)
    foreach ($p in $toRemove) { [void]$p.ParentNode.RemoveChild($p) }

    # ---------- 2. Латинская A → кириллическая А ----------
    # посимвольная замена с учётом того, что текст разбит по нескольким w:t
    $cyr = [string][char]1040
    $fixed = 0
    foreach ($p in @($body.SelectNodes('.//w:p', $ns))) {
        $ts = @($p.SelectNodes('.//w:t', $ns))
        if ($ts.Count -eq 0) { continue }
        $nodes = New-Object System.Collections.ArrayList
        $full = New-Object System.Text.StringBuilder
        foreach ($t in $ts) {
            [void]$nodes.Add(@{ Node = $t; Start = $full.Length; Len = $t.InnerText.Length })
            [void]$full.Append($t.InnerText)
        }
        $text = $full.ToString()
        if ($text -notmatch 'A') { continue }

        $hasApp = ($text -match 'ПРИЛОЖЕНИЕ')
        if (-not $hasApp -and $text -notmatch 'Рисунок') { continue }

        $hit = New-Object System.Collections.ArrayList
        for ($i = 0; $i -lt $text.Length; $i++) {
            if ($text[$i] -ne 'A') { continue }
            $before = if ($i -gt 0) { $text.Substring(0, $i) } else { '' }
            $after = if ($i -lt $text.Length - 1) { $text.Substring($i + 1) } else { '' }
            $isApp = ($before -match 'ПРИЛОЖЕНИЕ\s*$')
            $isFig = (($before -match 'Рисунок\s*$') -and ($after -match '^\.1'))
            $isWord = ($hasApp -and ($before -match '[^\p{L}\d]\z' -or $before -eq '') -and ($after -match '^[^\p{L}\d]' -or $after -eq ''))
            if ($isApp -or $isFig -or $isWord) { [void]$hit.Add($i) }
        }
        if ($hit.Count -eq 0) { continue }

        foreach ($pos in $hit) {
            foreach ($n in $nodes) {
                if ($pos -ge $n.Start -and $pos -lt ($n.Start + $n.Len)) {
                    $v = $n.Node.InnerText
                    $off = $pos - $n.Start
                    $n.Node.InnerText = $v.Substring(0, $off) + $cyr + $v.Substring($off + 1)
                    break
                }
            }
        }
        $fixed++
        Write-Output ('  исправлено: [' + (Get-PText $p).Trim() + ']')
    }
    Write-Output ('Исправлено абзацев: ' + $fixed)

    # ---------- 3. Подпись рис. 3 — в отдельный абзац ----------
    # находим образец pPr у существующей подписи рисунка (например "Рисунок 2 —")
    $samplePPr = $null
    foreach ($p in @($body.SelectNodes('w:p', $ns))) {
        $t = (Get-PText $p).Trim()
        if ($t -match '^Рисунок\s+\d+\s+[\u2014\u2013-]') {
            $pp = $p.SelectSingleNode('w:pPr', $ns)
            if ($null -ne $pp) { $samplePPr = $pp }
            break
        }
    }

    $done = $false
    foreach ($p in @($body.SelectNodes('w:p', $ns))) {
        if ($p.SelectNodes('.//w:drawing', $ns).Count -eq 0) { continue }
        # индекс run-а с рисунком среди дочерних узлов абзаца
        $drawIdx = -1
        $kids = @($p.ChildNodes)
        for ($k = 0; $k -lt $kids.Count; $k++) {
            if ($kids[$k].LocalName -eq 'r' -and $kids[$k].SelectNodes('.//w:drawing', $ns).Count -gt 0) { $drawIdx = $k; break }
        }
        if ($drawIdx -lt 0) { continue }

        # собираем узлы после run-а с рисунком, содержащие подпись
        $move = New-Object System.Collections.ArrayList
        $capText = ''
        for ($k = $drawIdx + 1; $k -lt $kids.Count; $k++) {
            $kn = $kids[$k]
            if ($kn.LocalName -notin @('r', 'proofErr', 'bookmarkStart', 'bookmarkEnd')) { continue }
            $kt = ''
            foreach ($t in @($kn.SelectNodes('.//w:t', $ns))) { $kt += $t.InnerText }
            if ($kn.LocalName -eq 'r' -and $kt.Trim() -eq '' -and $kn.SelectNodes('.//w:drawing', $ns).Count -eq 0 -and $kn.SelectNodes('.//w:br', $ns).Count -gt 0) { continue }
            [void]$move.Add($kn)
            $capText += $kt
        }
        if ($move.Count -eq 0) { continue }
        if ($capText -notmatch 'Рисунок\s*3') { continue }

        $newP = $xml.CreateElement('w', 'p', $W)
        if ($null -ne $samplePPr) { [void]$newP.AppendChild($samplePPr.CloneNode($true)) }
        else {
            $pPr = $p.SelectSingleNode('w:pPr', $ns)
            if ($null -ne $pPr) { [void]$newP.AppendChild($pPr.CloneNode($true)) }
        }
        foreach ($node in $move) {
            [void]$p.RemoveChild($node)
            [void]$newP.AppendChild($node)
        }
        [void]$p.ParentNode.InsertAfter($newP, $p)
        $done = $true
        Write-Output ('  подпись рис. 3 вынесена: [' + $capText.Trim() + ']')
        break
    }
    Write-Output ('Рис.3 вынесен: ' + $done)

    # ---------- сохранение в zip ----------
    $settings = New-Object System.Xml.XmlWriterSettings
    $settings.Encoding = New-Object System.Text.UTF8Encoding($false)
    $settings.Indent = $false
    $settings.OmitXmlDeclaration = $false
    $stream2 = $entry.Open()
    $stream2.SetLength(0)
    $writer = [System.Xml.XmlWriter]::Create($stream2, $settings)
    $xml.Save($writer)
    $writer.Flush(); $writer.Close()
    $stream2.Close()
    Write-Output 'XML записан в word/document.xml'
}
finally {
    $zip.Dispose()
}
Write-Output 'OK'
