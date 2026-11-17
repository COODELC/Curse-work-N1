# Функции правки .docx (OOXML) без Word: абзацы, таблицы, изображения.
# Подключается через . (dot-source) в скриптах раздела 2.

Add-Type -AssemblyName System.IO.Compression.FileSystem
Add-Type -AssemblyName System.Drawing

$script:ParaRx = [regex]'(?s)<w:p\b(?:(?!</w:p>).)*?</w:p>|<w:p\b[^>]*/>'

function Get-ParaText {
    param([string]$Xml)
    $sb = New-Object System.Text.StringBuilder
    foreach ($m in [regex]::Matches($Xml, '<w:t[^>]*>([^<]*)</w:t>')) { [void]$sb.Append($m.Groups[1].Value) }
    return $sb.ToString()
}

function Get-ParaMatch {
    param(
        [string]$Xml,
        [string]$Match,
        [int]$Occurrence = 1,
        [switch]$Exact
    )
    $hits = @()
    foreach ($m in $script:ParaRx.Matches($Xml)) {
        $t = Get-ParaText $m.Value
        $ok = if ($Exact) { $t.Trim() -eq $Match } else { $t.Contains($Match) }
        if ($ok) { $hits += $m }
    }
    if ($hits.Count -lt $Occurrence) { return $null }
    return $hits[$Occurrence - 1]
}

function Escape-XmlText {
    param([string]$Text)
    return $Text.Replace('&', '&amp;').Replace('<', '&lt;').Replace('>', '&gt;')
}

# Заменяет текст абзаца, сохраняя его форматирование (rPr первой строки).
function Set-ParaText {
    param([string]$Xml, [string]$Match, [string]$NewText, [int]$Occurrence = 1)

    $m = Get-ParaMatch -Xml $Xml -Match $Match -Occurrence $Occurrence
    if ($null -eq $m) { throw "Абзац не найден: '$Match'" }

    $p = $m.Value
    $script:__cnt = 0

    $p = [regex]::Replace($p, '(?s)<w:t(\s[^>]*)?>(.*?)</w:t>', {
        param($x)
        if ($script:__cnt -eq 0) {
            $script:__cnt = 1
            return '<w:t' + $x.Groups[1].Value + '>' + (Escape-XmlText $NewText) + '</w:t>'
        }
        return ''
    })

    # Пустые прогоны (без текста) убираем
    $p = [regex]::Replace($p, '(?s)<w:r>(?:(?!</w:r>).)*?</w:r>', {
        param($x)
        if ($x.Value -match '<w:t') { return $x.Value } else { return '' }
    })

    return $Xml.Remove($m.Index, $m.Length).Insert($m.Index, $p)
}

# Удаляет абзац целиком.
function Remove-Para {
    param([string]$Xml, [string]$Match, [int]$Occurrence = 1)

    $m = Get-ParaMatch -Xml $Xml -Match $Match -Occurrence $Occurrence
    if ($null -eq $m) { throw "Абзац не найден: '$Match'" }

    return $Xml.Remove($m.Index, $m.Length)
}

# Вставляет новый абзац с текстом сразу после абзаца-маркера.
# Шаблон оформления берётся у абзаца-образца (по умолчанию — подпись рисунка).
function Insert-ParaAfter {
    param(
        [string]$Xml,
        [string]$AfterMatch,
        [string]$NewText,
        [string]$TemplateMatch = $null,
        [int]$Occurrence = 1,
        [int]$TemplateOccurrence = 1,
        [switch]$Before,
        [switch]$Exact
    )

    $anchor = Get-ParaMatch -Xml $Xml -Match $AfterMatch -Occurrence $Occurrence -Exact:$Exact
    if ($null -eq $anchor) { throw "Абзац-маркер не найден: '$AfterMatch'" }

    if ($null -eq $TemplateMatch) {
        $tpl = Get-ParaMatch -Xml $Xml -Match 'Рисунок'
        if ($null -eq $tpl) { $tpl = Get-ParaMatch -Xml $Xml -Match $AfterMatch }
    }
    else {
        $tpl = Get-ParaMatch -Xml $Xml -Match $TemplateMatch -Occurrence $TemplateOccurrence
        if ($null -eq $tpl) { throw "Шаблон абзаца не найден: '$TemplateMatch'" }
    }

    $p = $tpl.Value
    $script:__cnt2 = 0
    $p = [regex]::Replace($p, '(?s)<w:t(\s[^>]*)?>(.*?)</w:t>', {
        param($x)
        if ($script:__cnt2 -eq 0) {
            $script:__cnt2 = 1
            return '<w:t' + $x.Groups[1].Value + '>' + (Escape-XmlText $NewText) + '</w:t>'
        }
        return ''
    })
    $p = [regex]::Replace($p, '(?s)<w:r>(?:(?!</w:r>).)*?</w:r>', {
        param($x)
        if ($x.Value -match '<w:t') { return $x.Value } else { return '' }
    })

    $pos = if ($Before) { $anchor.Index } else { $anchor.Index + $anchor.Length }
    return $Xml.Insert($pos, $p)
}

# Пересобирает таблицу, следующую за абзацем-подписью $Caption.
# $Rows — массив массивов строк; первая строка — заголовок.
# $TemplateCaption — подпись таблицы, шаблон которой берётся за образец (нужна,
# если у новой таблицы другое число колонок).
function Set-TableRows {
    param(
        [string]$Xml,
        [string]$Caption,
        [object[]]$Rows,
        [string]$TemplateCaption = $null,
        [int]$Occurrence = 1
    )

    $cap = Get-ParaMatch -Xml $Xml -Match $Caption -Occurrence $Occurrence
    if ($null -eq $cap) { throw "Подпись таблицы не найдена: '$Caption'" }

    $tblStart = $Xml.IndexOf('<w:tbl>', $cap.Index + $cap.Length)
    if ($tblStart -lt 0) { throw "Таблица после подписи '$Caption' не найдена" }

    # Конец таблицы ищем с учётом вложенных таблиц
    $depth = 0
    $pos = $tblStart
    $tblEnd = -1
    while ($pos -lt $Xml.Length) {
        $open = $Xml.IndexOf('<w:tbl>', $pos)
        $close = $Xml.IndexOf('</w:tbl>', $pos)
        if ($close -lt 0) { break }
        if ($open -ge 0 -and $open -lt $close) {
            $depth++
            $pos = $open + 7
            continue
        }
        $depth--
        $pos = $close + 8
        if ($depth -eq 0) { $tblEnd = $pos; break }
    }
    if ($tblEnd -lt 0) { throw "Таблица не закрыта" }

    $tbl = $Xml.Substring($tblStart, $tblEnd - $tblStart)

    # Шаблон: сама таблица (число колонок) или другая таблица-образец
    $tpl = $tbl
    if ($null -ne $TemplateCaption) {
        $cap2 = Get-ParaMatch -Xml $Xml -Match $TemplateCaption
        $s2 = $Xml.IndexOf('<w:tbl>', $cap2.Index + $cap2.Length)
        $depth2 = 0
        $pos2 = $s2
        $e2 = -1
        while ($pos2 -lt $Xml.Length) {
            $open2 = $Xml.IndexOf('<w:tbl>', $pos2)
            $close2 = $Xml.IndexOf('</w:tbl>', $pos2)
            if ($close2 -lt 0) { break }
            if ($open2 -ge 0 -and $open2 -lt $close2) { $depth2++; $pos2 = $open2 + 7; continue }
            $depth2--
            $pos2 = $close2 + 8
            if ($depth2 -eq 0) { $e2 = $pos2; break }
        }
        $tpl = $Xml.Substring($s2, $e2 - $s2)
    }

    $trRx = [regex]'(?s)<w:tr\b(?:(?!</w:tr>).)*?</w:tr>'
    $trs = $trRx.Matches($tpl)
    if ($trs.Count -lt 2) { throw "В шаблоне таблицы меньше двух строк" }

    $headTpl = $trs[0].Value
    $dataTpl = $trs[1].Value

    function Fill-Row {
        param([string]$RowXml, [string[]]$Values)

        $cells = [regex]::Matches($RowXml, '(?s)<w:tc>.*?</w:tc>')
        $out = $RowXml

        # Ячейки заменяем с конца, чтобы индексы не смещались
        for ($i = $cells.Count - 1; $i -ge 0; $i--) {
            $cell = $cells[$i].Value
            $val = if ($i -lt $Values.Count) { [string]$Values[$i] } else { '' }
            $first = $true

            $newCell = [regex]::Replace($cell, '(?s)<w:t(\s[^>]*)?>(.*?)</w:t>', {
                param($x)
                if ($first) {
                    $first = $false
                    return '<w:t' + $x.Groups[1].Value + '>' + (Escape-XmlText $val) + '</w:t>'
                }
                return ''
            })

            # Рисунки из строки-шаблона в новых строках не нужны:
            # их копирование даёт дубли идентификаторов и ломает документ.
            $newCell = [regex]::Replace($newCell, '(?s)<w:drawing>.*?</w:drawing>', '')

            $out = $out.Remove($cells[$i].Index, $cells[$i].Length).Insert($cells[$i].Index, $newCell)
        }

        return $out
    }

    # Каждая строка строится из шаблона: первая — из шапки, остальные — из данных
    $built = New-Object System.Collections.ArrayList
    for ($r = 0; $r -lt $Rows.Count; $r++) {
        $tplRow = if ($r -eq 0) { $headTpl } else { $dataTpl }
        $rowXml = Fill-Row -RowXml $tplRow -Values $Rows[$r]
        [void]$built.Add($rowXml)
    }

    # Заменяем все строки исходной таблицы на новые
    $trsOrig = $trRx.Matches($tbl)
    $first = $trsOrig[0].Index
    $last = $trsOrig[$trsOrig.Count - 1].Index + $trsOrig[$trsOrig.Count - 1].Length
    $newTbl = $tbl.Remove($first, $last - $first).Insert($first, ($built -join ''))

    return $Xml.Remove($tblStart, $tblEnd - $tblStart).Insert($tblStart, $newTbl)
}

# Заменяет картинку в документе: содержимое media и размер в EMU.
function Replace-DocImage {
    param(
        [System.IO.Compression.ZipArchive]$Zip,
        [ref]$XmlRef,
        [string]$MediaName,
        [string]$FilePath,
        [double]$MaxWidthIn = 6.5
    )

    $xml = $XmlRef.Value

    # rId этой картинки
    $relsEntry = $Zip.GetEntry('word/_rels/document.xml.rels')
    $sr = New-Object System.IO.StreamReader($relsEntry.Open(), [Text.Encoding]::UTF8)
    $rels = $sr.ReadToEnd(); $sr.Close()

    $relMatch = [regex]::Match($rels, ('Id="(rId\d+)"[^>]*Target="media/' + [regex]::Escape($MediaName) + '"'))
    if (-not $relMatch.Success) { $relMatch = [regex]::Match($rels, ('Target="media/' + [regex]::Escape($MediaName) + '"[^>]*Id="(rId\d+)"')) }
    if (-not $relMatch.Success) { throw "Не найдена связь для $MediaName" }
    $rid = $relMatch.Groups[1].Value

    $extRx = [regex]('(?s)<wp:extent cx="(\d+)" cy="(\d+)"[^>]*/>(?:(?!</w:drawing>).)*?r:embed="' + $rid + '"')
    $m = [regex]::Match($xml, $extRx)
    if (-not $m.Success) { throw "Не найден wp:extent для $rid ($MediaName)" }

    $oldCx = [int64]$m.Groups[1].Value
    $oldCy = [int64]$m.Groups[2].Value

    $img = [System.Drawing.Image]::FromFile($FilePath)
    $ratio = $img.Width / $img.Height
    $img.Dispose()

    $maxCx = [int64]([Math]::Round($MaxWidthIn * 914400))
    $newCx = if ($oldCx -le $maxCx) { $oldCx } else { $maxCx }
    $newCy = [int64][Math]::Round($newCx / $ratio)

    $pair = 'cx="' + $oldCx + '" cy="' + $oldCy + '"'
    $newPair = 'cx="' + $newCx + '" cy="' + $newCy + '"'
    $xml = $xml.Replace($pair, $newPair)
    $XmlRef.Value = $xml

    # файл картинки
    $entry = $Zip.GetEntry("word/media/$MediaName")
    if ($null -eq $entry) { throw "Нет записи word/media/$MediaName" }
    $entry.Delete()
    $newEntry = $Zip.CreateEntry("word/media/$MediaName", [System.IO.Compression.CompressionLevel]::Optimal)
    $in = [System.IO.File]::OpenRead($FilePath)
    $out = $newEntry.Open()
    $in.CopyTo($out)
    $out.Close(); $in.Close()

    return "OK $MediaName -> $(Split-Path $FilePath -Leaf): $newCx x $newCy EMU"
}

# Вставляет новую картинку: копирует абзац-образец с рисунком, подменяет
# содержимое файла, rId и размер. За абзацем-маркером добавляется подпись.
function Insert-DocImage {
    param(
        [System.IO.Compression.ZipArchive]$Zip,
        [ref]$XmlRef,
        [string]$AfterMatch,
        [string]$FilePath,
        [string]$CaptionText,
        [string]$TemplateImageMatch,
        [double]$WidthIn = 6.0
    )

    $xml = $XmlRef.Value

    # Образец: абзац с рисунком стоит перед своей подписью.
    # TemplateImageMatch — текст подписи-образца.
    $capTpl = Get-ParaMatch -Xml $xml -Match $TemplateImageMatch
    if ($null -eq $capTpl) { throw "Не найдена подпись-образец: '$TemplateImageMatch'" }

    $tpl = $null
    foreach ($mm in $script:ParaRx.Matches($xml)) {
        if ($mm.Index + $mm.Length -le $capTpl.Index -and $mm.Value -match '<w:drawing') { $tpl = $mm }
    }
    if ($null -eq $tpl) { throw "Не найден абзац с рисунком перед подписью '$TemplateImageMatch'" }

    # Новое имя файла и связь
    $existing = @($Zip.Entries | Where-Object { $_.FullName -like 'word/media/image*.png' } | ForEach-Object {
        [int]([regex]::Match($_.FullName, 'image(\d+)\.png').Groups[1].Value)
    })
    $newNum = 1 + ($existing | Measure-Object -Maximum).Maximum
    $newMedia = "image$newNum.png"

    $relsEntry = $Zip.GetEntry('word/_rels/document.xml.rels')
    $sr = New-Object System.IO.StreamReader($relsEntry.Open(), [Text.Encoding]::UTF8)
    $rels = $sr.ReadToEnd(); $sr.Close()

    $relIds = [regex]::Matches($rels, 'Id="rId(\d+)"') | ForEach-Object { [int]$_.Groups[1].Value }
    $newRelId = 'rId' + (1 + ($relIds | Measure-Object -Maximum).Maximum)

    $relXml = '<Relationship Id="' + $newRelId + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/' + $newMedia + '"/>'
    $rels = $rels.Replace('</Relationships>', $relXml + '</Relationships>')

    $relsEntry.Delete()
    $rels2 = $Zip.CreateEntry('word/_rels/document.xml.rels', [System.IO.Compression.CompressionLevel]::Optimal)
    $sw = New-Object System.IO.StreamWriter($rels2.Open(), (New-Object Text.UTF8Encoding($false)))
    $sw.Write($rels); $sw.Close()

    # файл картинки
    $newEntry = $Zip.CreateEntry("word/media/$newMedia", [System.IO.Compression.CompressionLevel]::Optimal)
    $in = [System.IO.File]::OpenRead($FilePath)
    $out = $newEntry.Open()
    $in.CopyTo($out)
    $out.Close(); $in.Close()

    # абзац с рисунком: копия образца с новым rId и размером
    $img = [System.Drawing.Image]::FromFile($FilePath)
    $ratio = $img.Width / $img.Height
    $img.Dispose()

    $p = $tpl.Value

    # В образце может быть несколько объектов (рисунок, фигура, OLE).
    # Оставляем только первый рисунок, остальные удаляем.
    $drawings = [regex]::Matches($p, '(?s)<w:drawing>.*?</w:drawing>')
    if ($drawings.Count -gt 1) {
        $p = $p.Remove($drawings[1].Index, $drawings[1].Length)
        for ($i = 2; $i -lt $drawings.Count; $i++) {
            $p = $p.Remove($drawings[$i].Index, $drawings[$i].Length)
        }
    }

    # Кадрирование исходного изображения не переносим на новый рисунок
    $p = [regex]::Replace($p, '(?s)<a:srcRect[^>]*/>', '')

    $m = [regex]::Match($p, '(?s)<wp:extent cx="(\d+)" cy="(\d+)"')
    $oldCx = [int64]$m.Groups[1].Value
    $oldCy = [int64]$m.Groups[2].Value
    $newCx = [int64][Math]::Round($WidthIn * 914400)
    $newCy = [int64][Math]::Round($newCx / $ratio)

    $p = $p.Replace('cx="' + $oldCx + '" cy="' + $oldCy + '"', 'cx="' + $newCx + '" cy="' + $newCy + '"')
    $p = [regex]::Replace($p, 'r:embed="rId\d+"', ('r:embed="' + $newRelId + '"'))
    $p = [regex]::Replace($p, '<wp:docPr id="\d+" name="[^"]*"', ('<wp:docPr id="' + (900 + (Get-Random -Maximum 90)) + '" name="Рисунок"'))
    $p = [regex]::Replace($p, '<a:ext cx="\d+" cy="\d+"/>', ('<a:ext cx="' + $newCx + '" cy="' + $newCy + '"/>'))

    $anchor = Get-ParaMatch -Xml $xml -Match $AfterMatch
    if ($null -eq $anchor) { throw "Абзац-маркер не найден: '$AfterMatch'" }

    $pos = $anchor.Index + $anchor.Length
    $xml = $xml.Insert($pos, $p)

    # Подпись рисунка сразу после вставленного рисунка
    if (-not [string]::IsNullOrWhiteSpace($CaptionText)) {
        $capXml = $capTpl.Value
        $script:__cnt3 = 0
        $capXml = [regex]::Replace($capXml, '(?s)<w:t(\s[^>]*)?>(.*?)</w:t>', {
            param($x)
            if ($script:__cnt3 -eq 0) {
                $script:__cnt3 = 1
                return '<w:t' + $x.Groups[1].Value + '>' + (Escape-XmlText $CaptionText) + '</w:t>'
            }
            return ''
        })
        $capXml = [regex]::Replace($capXml, '(?s)<w:r>(?:(?!</w:r>).)*?</w:r>', {
            param($x)
            if ($x.Value -match '<w:t') { return $x.Value } else { return '' }
        })

        $pos2 = $pos + $p.Length
        $xml = $xml.Insert($pos2, $capXml)
    }

    $XmlRef.Value = $xml

    return "OK вставлено $newMedia ($newRelId): $newCx x $newCy EMU"
}

# Возвращает абзац, который находился внутри последней ячейки таблицы,
# сразу после самой таблицы (так восстанавливается заголовок раздела).
function Insert-ParaAfterTable {
    param(
        [string]$Xml,
        [string]$Caption,
        [string]$ParaXml
    )

    $cap = Get-ParaMatch -Xml $Xml -Match $Caption
    if ($null -eq $cap) { throw "Подпись таблицы не найдена: '$Caption'" }

    $tblStart = $Xml.IndexOf('<w:tbl>', $cap.Index + $cap.Length)
    if ($tblStart -lt 0) { throw "Таблица после подписи '$Caption' не найдена" }

    $depth = 0
    $pos = $tblStart
    $tblEnd = -1
    while ($pos -lt $Xml.Length) {
        $open = $Xml.IndexOf('<w:tbl>', $pos)
        $close = $Xml.IndexOf('</w:tbl>', $pos)
        if ($close -lt 0) { break }
        if ($open -ge 0 -and $open -lt $close) { $depth++; $pos = $open + 7; continue }
        $depth--
        $pos = $close + 8
        if ($depth -eq 0) { $tblEnd = $pos; break }
    }
    if ($tblEnd -lt 0) { throw "Таблица не закрыта" }

    return $Xml.Insert($tblEnd, $ParaXml)
}

# Сохраняет document.xml обратно в архив.
function Save-DocumentXml {
    param([System.IO.Compression.ZipArchive]$Zip, [string]$Xml)
    $entry = $Zip.GetEntry('word/document.xml')
    $entry.Delete()
    $newEntry = $Zip.CreateEntry('word/document.xml', [System.IO.Compression.CompressionLevel]::Optimal)
    $sw = New-Object System.IO.StreamWriter($newEntry.Open(), (New-Object Text.UTF8Encoding($false)))
    $sw.Write($Xml); $sw.Close()
}
