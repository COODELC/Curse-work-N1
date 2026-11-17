# Правка .docx через XML DOM (OOXML) без Word.
# Надёжнее regex-подходов: корректная работа с w:drawing/w:object, уникальные id рисунков.

Add-Type -AssemblyName System.IO.Compression.FileSystem
Add-Type -AssemblyName System.Drawing

$script:W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
$script:R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
$script:PKG_REL = 'http://schemas.openxmlformats.org/package/2006/relationships'

function New-DocxContext {
    param([Parameter(Mandatory = $true)][string]$Path)

    $zip = [System.IO.Compression.ZipFile]::Open($Path, 'Update')
    $entry = $zip.GetEntry('word/document.xml')
    $sr = New-Object System.IO.StreamReader($entry.Open(), [Text.Encoding]::UTF8)
    $xmlText = $sr.ReadToEnd()
    $sr.Close()

    $doc = New-Object System.Xml.XmlDocument
    $doc.PreserveWhitespace = $true
    $doc.LoadXml($xmlText)

    $nsm = New-Object System.Xml.XmlNamespaceManager($doc.NameTable)
    $nsm.AddNamespace('w', $script:W)
    $nsm.AddNamespace('r', $script:R)

    return [pscustomobject]@{
        Path   = $Path
        Zip    = $zip
        Doc    = $doc
        Nsm    = $nsm
        RelsDoc = $null
        RelsText = $null
        MaxDocPrId = 0
        ScopeStart = $null
    }
}

# Задаёт начало области поиска: абзацы до него (например, оглавление)
# при поиске по тексту не рассматриваются.
function Set-DocxScope {
    param($Ctx, [string]$AnchorMatch)

    $blocks = Get-BodyBlocks $Ctx
    for ($i = 0; $i -lt $blocks.Count; $i++) {
        $t = (Get-ParaTextOf $blocks[$i]).Trim()
        if ($t -eq $AnchorMatch) {
            $Ctx.ScopeStart = $i
            Write-Output ("OK область правки с блока " + $i + ": " + $AnchorMatch)
            return
        }
    }
    throw ("Не найдено начало области: '" + $AnchorMatch + "'")
}

function Save-DocxContext {
    param($Ctx)

    # ВАЖНО: записи перезаписываются "по месту" (SetLength(0)) — удаление и
    # пересоздание записи в ZipArchive ломает пакет, и Word отказывается
    # открывать документ ("Файл повреждён").
    if ($null -ne $Ctx.RelsText) {
        $relsEntry = $Ctx.Zip.GetEntry('word/_rels/document.xml.rels')
        $s = $relsEntry.Open()
        $s.SetLength(0)
        $sw = New-Object System.IO.StreamWriter($s, (New-Object Text.UTF8Encoding($false)))
        $sw.Write($Ctx.RelsText); $sw.Close(); $s.Close()
    }

    $xmlOut = $Ctx.Doc.OuterXml
    $entry = $Ctx.Zip.GetEntry('word/document.xml')
    $s = $entry.Open()
    $s.SetLength(0)
    $sw = New-Object System.IO.StreamWriter($s, (New-Object Text.UTF8Encoding($false)))
    $sw.Write($xmlOut); $sw.Close(); $s.Close()
}

function Close-DocxContext {
    param($Ctx)
    if ($null -ne $Ctx) { $Ctx.Zip.Dispose() }
}

# ---------- элементы ----------

function Get-BodyBlocks {
    param($Ctx)
    $body = $Ctx.Doc.SelectSingleNode('//w:body', $Ctx.Nsm)
    $list = New-Object System.Collections.ArrayList
    foreach ($n in $body.ChildNodes) {
        if ($n.NodeType -eq 'Element' -and ($n.LocalName -eq 'p' -or $n.LocalName -eq 'tbl')) {
            [void]$list.Add($n)
        }
    }
    return $list
}

function Get-ParaTextOf {
    param($p)
    $ns = New-Object System.Xml.XmlNamespaceManager($p.OwnerDocument.NameTable)
    $ns.AddNamespace('w', $script:W)
    $sb = New-Object System.Text.StringBuilder
    foreach ($t in $p.SelectNodes('.//w:t', $ns)) { [void]$sb.Append($t.InnerText) }
    return $sb.ToString()
}

function Find-Para {
    param(
        $Ctx,
        [string]$Match,
        [int]$Occurrence = 1,
        [switch]$Exact,
        [switch]$BodyLevel
    )
    $hits = New-Object System.Collections.ArrayList
    if ($BodyLevel) {
        $cands = Get-BodyBlocks $Ctx
        if ($null -ne $Ctx.ScopeStart) {
            $cands = @($cands | Select-Object -Skip $Ctx.ScopeStart)
        }
    }
    else {
        $cands = $Ctx.Doc.SelectNodes('//w:p', $Ctx.Nsm)
    }
    foreach ($p in $cands) {
        $t = (Get-ParaTextOf $p).Trim()
        $ok = if ($Exact) { $t -eq $Match } else { $t.Contains($Match) }
        if ($ok) { [void]$hits.Add($p) }
    }
    if ($hits.Count -lt $Occurrence) { return $null }
    return $hits[$Occurrence - 1]
}

function Set-ParaText {
    param($p, [string]$Text)

    $ns = New-Object System.Xml.XmlNamespaceManager($p.OwnerDocument.NameTable)
    $ns.AddNamespace('w', $script:W)

    $runs = @($p.SelectNodes('./w:r', $ns))
    $target = $null
    foreach ($r in $runs) {
        if ($r.SelectSingleNode('./w:t', $ns)) { $target = $r; break }
    }
    if ($null -eq $target) {
        # создаём новый run с форматированием первого существующего
        if ($runs.Count -gt 0) {
            $target = $runs[0].CloneNode($true)
            $p.AppendChild($target) | Out-Null
        }
        else {
            $target = $p.OwnerDocument.CreateElement('w', 'r', $script:W)
            $p.AppendChild($target) | Out-Null
        }
    }

    # убираем все w:t во всех прогонах, кроме первого найденного в target
    $firstT = $target.SelectSingleNode('./w:t', $ns)
    foreach ($r in $runs) {
        foreach ($t in @($r.SelectNodes('./w:t', $ns))) {
            if ($t -eq $firstT) { continue }
            $r.RemoveChild($t) | Out-Null
        }
    }

    if ($null -eq $firstT) {
        $firstT = $p.OwnerDocument.CreateElement('w', 't', $script:W)
        $target.AppendChild($firstT) | Out-Null
    }
    $firstT.InnerText = $Text
    # ВАЖНО: xml:space создаём с префиксом xml явно — иначе .NET сериализует
    # атрибут как d5p1:space, и Word считает документ повреждённым.
    $xmlNs = 'http://www.w3.org/XML/1998/namespace'
    if ([string]::IsNullOrEmpty($firstT.GetAttribute('space', $xmlNs))) {
        $attr = $p.OwnerDocument.CreateAttribute('xml', 'space', $xmlNs)
        $attr.Value = 'preserve'
        $firstT.Attributes.Append($attr) | Out-Null
    }
    return $p
}

function Remove-ParaBy {
    param($Ctx, [string]$Match, [int]$Occurrence = 1, [switch]$Exact)
    $p = Find-Para -Ctx $Ctx -Match $Match -Occurrence $Occurrence -Exact:$Exact
    if ($null -eq $p) { throw ("Абзац не найден: '" + $Match + "'") }
    $p.ParentNode.RemoveChild($p) | Out-Null
    Write-Output ("OK удалён абзац: " + $Match)
}

# ---------- таблицы ----------

function Get-TableAfter {
    param($Ctx, $p)
    $node = $p.NextSibling
    while ($null -ne $node) {
        if ($node.NodeType -eq 'Element' -and $node.LocalName -eq 'tbl') { return $node }
        $node = $node.NextSibling
    }
    return $null
}

function Find-TableByCaption {
    param($Ctx, [string]$Caption, [int]$Occurrence = 1)
    $blocks = Get-BodyBlocks $Ctx
    if ($null -ne $Ctx.ScopeStart) {
        $blocks = @($blocks | Select-Object -Skip $Ctx.ScopeStart)
    }
    $hit = 0
    foreach ($b in $blocks) {
        if ($b.LocalName -ne 'p') { continue }
        $t = (Get-ParaTextOf $b).Trim()
        if ($t.Contains($Caption)) {
            $hit++
            if ($hit -eq $Occurrence) {
                $tbl = Get-TableAfter -Ctx $Ctx -p $b
                if ($null -eq $tbl) { throw ("Таблица после подписи '" + $Caption + "' не найдена") }
                return $tbl
            }
        }
    }
    throw ("Подпись таблицы не найдена: '" + $Caption + "'")
}

function Clear-CellContent {
    param($tc)
    $ns = New-Object System.Xml.XmlNamespaceManager($tc.OwnerDocument.NameTable)
    $ns.AddNamespace('w', $script:W)
    $ps = @($tc.SelectNodes('./w:p', $ns))
    if ($ps.Count -eq 0) {
        $p = $tc.OwnerDocument.CreateElement('w', 'p', $script:W)
        $tc.AppendChild($p) | Out-Null
        return
    }
    # оставляем первый абзац, остальные удаляем
    for ($i = 1; $i -lt $ps.Count; $i++) { $tc.RemoveChild($ps[$i]) | Out-Null }
    $p = $ps[0]
    # удаляем рисунки и OLE
    foreach ($d in @($p.SelectNodes('.//w:drawing', $ns))) { $d.ParentNode.RemoveChild($d) | Out-Null }
    foreach ($o in @($p.SelectNodes('.//w:object', $ns))) { $o.ParentNode.RemoveChild($o) | Out-Null }
}

function Set-TableRows {
    param(
        $Ctx,
        [string]$Caption,
        [object[]]$Rows,
        [string]$TemplateCaption = $null,
        [int]$Occurrence = 1
    )

    $tbl = Find-TableByCaption -Ctx $Ctx -Caption $Caption -Occurrence $Occurrence

    $ns = New-Object System.Xml.XmlNamespaceManager($Ctx.Doc.NameTable)
    $ns.AddNamespace('w', $script:W)

    # источник строк-шаблонов
    if ($null -ne $TemplateCaption) {
        $tplTbl = Find-TableByCaption -Ctx $Ctx -Caption $TemplateCaption
    }
    else {
        $tplTbl = $tbl
    }

    $tplRows = @($tplTbl.SelectNodes('./w:tr', $ns))
    if ($tplRows.Count -lt 2) { throw 'В шаблоне таблицы меньше двух строк' }

    $newRows = New-Object System.Collections.ArrayList
    for ($r = 0; $r -lt $Rows.Count; $r++) {
        $src = if ($r -eq 0) { $tplRows[0] } else { $tplRows[1] }
        $tr = $src.CloneNode($true)
        $cells = @($tr.SelectNodes('./w:tc', $ns))
        $vals = $Rows[$r]
        for ($i = 0; $i -lt $cells.Count; $i++) {
            $val = ''
            if ($i -lt $vals.Count) { $val = [string]$vals[$i] }
            Clear-CellContent -tc $cells[$i]
            $ps = @($cells[$i].SelectNodes('./w:p', $ns))
            Set-ParaText -p $ps[0] -Text $val | Out-Null
        }
        [void]$newRows.Add($tr)
    }

    # заменяем строки: новые вставляем перед первой старой, затем удаляем старые
    $oldRows = @($tbl.SelectNodes('./w:tr', $ns))
    $first = $oldRows[0]
    foreach ($tr in $newRows) { $tbl.InsertBefore($tr, $first) | Out-Null }
    foreach ($tr in $oldRows) { $tbl.RemoveChild($tr) | Out-Null }

    # сетка колонок — из шаблона, если другое число колонок
    if ($null -ne $TemplateCaption) {
        $tplGrid = $tplTbl.SelectSingleNode('./w:tblGrid', $ns)
        $myGrid = $tbl.SelectSingleNode('./w:tblGrid', $ns)
        if ($null -ne $tplGrid -and $null -ne $myGrid) {
            $myGrid.ParentNode.ReplaceChild($tplGrid.CloneNode($true), $myGrid) | Out-Null
        }
    }
    Write-Output ("OK таблица '" + $Caption + "' пересобрана: " + $Rows.Count + " строк")
}

# ---------- изображения ----------

function Get-Rels {
    param($Ctx)
    if ($null -ne $Ctx.RelsText) { return $Ctx.RelsText }
    $e = $Ctx.Zip.GetEntry('word/_rels/document.xml.rels')
    $sr = New-Object System.IO.StreamReader($e.Open(), [Text.Encoding]::UTF8)
    $Ctx.RelsText = $sr.ReadToEnd()
    $sr.Close()
    return $Ctx.RelsText
}

function Add-MediaFile {
    param($Ctx, [string]$FilePath)

    $rels = Get-Rels $Ctx
    $ids = [regex]::Matches($rels, 'Id="rId(\d+)"') | ForEach-Object { [int]$_.Groups[1].Value }
    $newRelId = 'rId' + (1 + ($ids | Measure-Object -Maximum).Maximum)

    $nums = @($Ctx.Zip.Entries | Where-Object { $_.FullName -like 'word/media/image*' } | ForEach-Object {
        $mm = [regex]::Match($_.FullName, 'image(\d+)')
        if ($mm.Success) { [int]$mm.Groups[1].Value } else { 0 }
    })
    $max = 0
    if ($nums.Count -gt 0) { $max = ($nums | Measure-Object -Maximum).Maximum }
    $newMedia = 'image' + ($max + 1) + '.png'

    $e = $Ctx.Zip.CreateEntry('word/media/' + $newMedia, [System.IO.Compression.CompressionLevel]::Optimal)
    $in = [System.IO.File]::OpenRead($FilePath)
    $out = $e.Open()
    $in.CopyTo($out)
    $out.Close(); $in.Close()

    $relXml = '<Relationship Id="' + $newRelId + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/' + $newMedia + '"/>'
    $Ctx.RelsText = $rels.Replace('</Relationships>', $relXml + '</Relationships>')

    return @{ RelId = $newRelId; Media = $newMedia }
}

function Get-MaxDocPrId {
    param($Ctx)
    $max = 0
    $ns = New-Object System.Xml.XmlNamespaceManager($Ctx.Doc.NameTable)
    $ns.AddNamespace('wp', 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing')
    foreach ($d in $Ctx.Doc.SelectNodes('//wp:docPr', $ns)) {
        $v = 0
        if ([int]::TryParse($d.GetAttribute('id'), [ref]$v)) { if ($v -gt $max) { $max = $v } }
    }
    return $max
}

function New-DrawingXml {
    param([string]$RelId, [string]$FileName, [int]$DocPrId, [int64]$Cx, [int64]$Cy)

    $a = 'http://schemas.openxmlformats.org/drawingml/2006/main'
    $pic = 'http://schemas.openxmlformats.org/drawingml/2006/picture'
    $wp = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'
    $r = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

    return @"
<w:drawing xmlns:w="$script:W" xmlns:wp="$wp" xmlns:a="$a" xmlns:pic="$pic" xmlns:r="$r"><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="$Cx" cy="$Cy"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:docPr id="$DocPrId" name="Рисунок $DocPrId"/><wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr><a:graphic><a:graphicData uri="$pic"><pic:pic><pic:nvPicPr><pic:cNvPr id="0" name="$FileName"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="$RelId"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="$Cx" cy="$Cy"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing>
"@
}

function Get-PngSize {
    param([string]$FilePath)
    $img = [System.Drawing.Image]::FromFile($FilePath)
    $w = $img.Width; $h = $img.Height
    $img.Dispose()
    return @{ W = $w; H = $h }
}

# Заменяет рисунок в абзаце (или в предшествующем абзаце с рисунком) новым PNG.
function Set-Figure {
    param(
        $Ctx,
        [string]$CaptionMatch,
        [string]$FilePath,
        [string]$NewCaption,
        [double]$WidthIn = 6.2,
        [switch]$CaptionHasDrawing,
        [switch]$NoCaptionChange
    )

    $cap = Find-Para -Ctx $Ctx -Match $CaptionMatch -BodyLevel
    if ($null -eq $cap) { throw ("Не найдена подпись рисунка: '" + $CaptionMatch + "'") }

    $ns = New-Object System.Xml.XmlNamespaceManager($Ctx.Doc.NameTable)
    $ns.AddNamespace('w', $script:W)

    # абзац с рисунком: либо сама подпись, либо ближайший предыдущий с рисунком
    $target = $null
    if ($CaptionHasDrawing -and $cap.SelectSingleNode('.//w:drawing', $ns)) {
        $target = $cap
    }
    else {
        $node = $cap.PreviousSibling
        $guard = 0
        while ($null -ne $node -and $guard -lt 4) {
            $guard++
            if ($node.NodeType -eq 'Element' -and $node.LocalName -eq 'p') {
                if ($node.SelectSingleNode('.//w:drawing', $ns) -or $node.SelectSingleNode('.//w:object', $ns)) {
                    $target = $node; break
                }
            }
            $node = $node.PreviousSibling
        }
    }
    if ($null -eq $target) { $target = $cap }

    # удаляем старые рисунки и OLE
    foreach ($d in @($target.SelectNodes('.//w:drawing', $ns))) { $d.ParentNode.RemoveChild($d) | Out-Null }
    foreach ($o in @($target.SelectNodes('.//w:object', $ns))) { $o.ParentNode.RemoveChild($o) | Out-Null }

    $size = Get-PngSize -FilePath $FilePath
    $cx = [int64][Math]::Round($WidthIn * 914400)
    $cy = [int64][Math]::Round($cx * $size.H / $size.W)

    $media = Add-MediaFile -Ctx $Ctx -FilePath $FilePath
    $Ctx.MaxDocPrId = (Get-MaxDocPrId $Ctx) + 1
    $drawingXml = New-DrawingXml -RelId $media.RelId -FileName $media.Media -DocPrId $Ctx.MaxDocPrId -Cx $cx -Cy $cy

    $frag = $Ctx.Doc.CreateDocumentFragment()
    $frag.InnerXml = '<w:r xmlns:w="' + $script:W + '">' + $drawingXml + '</w:r>'
    # если в абзаце уже есть текст (подпись в одном абзаце с рисунком) — рисунок ставим перед текстом
    $textRun = $null
    foreach ($r in @($target.SelectNodes('./w:r', $ns))) {
        if ($r.SelectSingleNode('./w:t', $ns)) { $textRun = $r; break }
    }
    if ($null -ne $textRun) {
        $target.InsertBefore($frag, $textRun) | Out-Null
    }
    else {
        $target.AppendChild($frag) | Out-Null
    }

    if (-not $NoCaptionChange -and -not [string]::IsNullOrWhiteSpace($NewCaption)) {
        Set-ParaText -p $cap -Text $NewCaption | Out-Null
    }

    Write-Output ("OK рисунок '" + $CaptionMatch + "' -> " + $media.Media + " (" + $cx + 'x' + $cy + ' EMU)')
}

# Вставляет новый рисунок с подписью после абзаца-маркера.
function Insert-Figure {
    param(
        $Ctx,
        [string]$AfterMatch,
        [string]$FilePath,
        [string]$CaptionText,
        [string]$CaptionTemplateMatch,
        [double]$WidthIn = 6.2
    )

    $anchor = Find-Para -Ctx $Ctx -Match $AfterMatch -BodyLevel
    if ($null -eq $anchor) { throw ("Абзац-маркер не найден: '" + $AfterMatch + "'") }

    $capTpl = Find-Para -Ctx $Ctx -Match $CaptionTemplateMatch -BodyLevel
    if ($null -eq $capTpl) { throw ("Не найден шаблон подписи: '" + $CaptionTemplateMatch + "'") }

    # новый абзац с рисунком на основе шаблона подписи (копия без текста)
    $imgPara = $capTpl.CloneNode($true)
    $ns = New-Object System.Xml.XmlNamespaceManager($Ctx.Doc.NameTable)
    $ns.AddNamespace('w', $script:W)
    foreach ($t in @($imgPara.SelectNodes('.//w:t', $ns))) { $t.ParentNode.RemoveChild($t) | Out-Null }
    foreach ($d in @($imgPara.SelectNodes('.//w:drawing', $ns))) { $d.ParentNode.RemoveChild($d) | Out-Null }
    foreach ($o in @($imgPara.SelectNodes('.//w:object', $ns))) { $o.ParentNode.RemoveChild($o) | Out-Null }

    $size = Get-PngSize -FilePath $FilePath
    $cx = [int64][Math]::Round($WidthIn * 914400)
    $cy = [int64][Math]::Round($cx * $size.H / $size.W)

    $media = Add-MediaFile -Ctx $Ctx -FilePath $FilePath
    $Ctx.MaxDocPrId = (Get-MaxDocPrId $Ctx) + 1
    $drawingXml = New-DrawingXml -RelId $media.RelId -FileName $media.Media -DocPrId $Ctx.MaxDocPrId -Cx $cx -Cy $cy
    $frag = $Ctx.Doc.CreateDocumentFragment()
    $frag.InnerXml = '<w:r xmlns:w="' + $script:W + '">' + $drawingXml + '</w:r>'
    $imgPara.AppendChild($frag) | Out-Null

    # абзац подписи
    $capPara = $capTpl.CloneNode($true)
    Set-ParaText -p $capPara -Text $CaptionText | Out-Null

    $anchor.ParentNode.InsertAfter($imgPara, $anchor) | Out-Null
    $anchor.ParentNode.InsertAfter($capPara, $imgPara) | Out-Null

    Write-Output ("OK вставлен рисунок " + $media.Media + " с подписью: " + $CaptionText)
}

# Вставляет абзац с текстом после (или перед) абзацем-маркером.
function Insert-Para {
    param(
        $Ctx,
        [string]$AfterMatch,
        [string]$NewText,
        [string]$TemplateMatch,
        [int]$Occurrence = 1,
        [int]$TemplateOccurrence = 1,
        [switch]$Before,
        [switch]$Exact
    )

    $anchor = Find-Para -Ctx $Ctx -Match $AfterMatch -Occurrence $Occurrence -Exact:$Exact -BodyLevel
    if ($null -eq $anchor) { throw ("Абзац-маркер не найден: '" + $AfterMatch + "'") }

    $tpl = Find-Para -Ctx $Ctx -Match $TemplateMatch -Occurrence $TemplateOccurrence -BodyLevel
    if ($null -eq $tpl) { throw ("Шаблон абзаца не найден: '" + $TemplateMatch + "'") }

    $p = $tpl.CloneNode($true)
    $ns = New-Object System.Xml.XmlNamespaceManager($Ctx.Doc.NameTable)
    $ns.AddNamespace('w', $script:W)
    foreach ($d in @($p.SelectNodes('.//w:drawing', $ns))) { $d.ParentNode.RemoveChild($d) | Out-Null }
    foreach ($o in @($p.SelectNodes('.//w:object', $ns))) { $o.ParentNode.RemoveChild($o) | Out-Null }
    Set-ParaText -p $p -Text $NewText | Out-Null

    if ($Before) {
        $anchor.ParentNode.InsertBefore($p, $anchor) | Out-Null
    }
    else {
        $anchor.ParentNode.InsertAfter($p, $anchor) | Out-Null
    }
    Write-Output ('OK вставлен абзац после: ' + $AfterMatch)
}

# Вставляет абзац с текстом перед указанным абзацем (по тексту).
function Insert-ParaBefore {
    param($Ctx, [string]$BeforeMatch, [string]$NewText, [string]$TemplateMatch, [switch]$Exact)
    Insert-Para -Ctx $Ctx -AfterMatch $BeforeMatch -NewText $NewText -TemplateMatch $TemplateMatch -Before -Exact:$Exact | Out-Null
}

# Абзацы-заглушки «ВАШ ТЕКСТ…» в порядке следования по документу.
function Get-PlaceholderParas {
    param($Ctx)
    $res = New-Object System.Collections.ArrayList
    $blocks = Get-BodyBlocks $Ctx
    if ($null -ne $Ctx.ScopeStart) {
        $blocks = @($blocks | Select-Object -Skip $Ctx.ScopeStart)
    }
    foreach ($b in $blocks) {
        if ($b.LocalName -ne 'p') { continue }
        $t = (Get-ParaTextOf $b).Trim()
        if ($t.StartsWith('ВАШ ТЕКСТ')) { [void]$res.Add($b) }
    }
    return $res
}

# Удаляет рисунки и OLE из абзаца-маркера.
function Remove-Figures {
    param($Ctx, [string]$ParaMatch, [int]$Occurrence = 1)
    $p = Find-Para -Ctx $Ctx -Match $ParaMatch -Occurrence $Occurrence -BodyLevel
    if ($null -eq $p) { throw ("Абзац-маркер не найден: '" + $ParaMatch + "'") }
    $ns = New-Object System.Xml.XmlNamespaceManager($Ctx.Doc.NameTable)
    $ns.AddNamespace('w', $script:W)
    foreach ($d in @($p.SelectNodes('.//w:drawing', $ns))) { $d.ParentNode.RemoveChild($d) | Out-Null }
    foreach ($o in @($p.SelectNodes('.//w:object', $ns))) { $o.ParentNode.RemoveChild($o) | Out-Null }
    Write-Output ("OK рисунки удалены из абзаца: " + $ParaMatch)
}

# Заменяет рисунок внутри абзаца (текст и рисунок в одном абзаце) и меняет текст.
function Set-FigureInPara {
    param(
        $Ctx,
        [string]$ParaMatch,
        [string]$FilePath,
        [string]$NewText,
        [double]$WidthIn = 6.2,
        [int]$Occurrence = 1
    )
    $p = Find-Para -Ctx $Ctx -Match $ParaMatch -Occurrence $Occurrence -BodyLevel
    if ($null -eq $p) { throw ("Абзац-маркер не найден: '" + $ParaMatch + "'") }

    $ns = New-Object System.Xml.XmlNamespaceManager($Ctx.Doc.NameTable)
    $ns.AddNamespace('w', $script:W)
    foreach ($d in @($p.SelectNodes('.//w:drawing', $ns))) { $d.ParentNode.RemoveChild($d) | Out-Null }
    foreach ($o in @($p.SelectNodes('.//w:object', $ns))) { $o.ParentNode.RemoveChild($o) | Out-Null }

    $size = Get-PngSize -FilePath $FilePath
    $cx = [int64][Math]::Round($WidthIn * 914400)
    $cy = [int64][Math]::Round($cx * $size.H / $size.W)

    $media = Add-MediaFile -Ctx $Ctx -FilePath $FilePath
    $Ctx.MaxDocPrId = (Get-MaxDocPrId $Ctx) + 1
    $drawingXml = New-DrawingXml -RelId $media.RelId -FileName $media.Media -DocPrId $Ctx.MaxDocPrId -Cx $cx -Cy $cy
    $frag = $Ctx.Doc.CreateDocumentFragment()
    $frag.InnerXml = '<w:r xmlns:w="' + $script:W + '">' + $drawingXml + '</w:r>'
    # рисунок ставим перед текстом абзаца
    $textRun = $null
    foreach ($r in @($p.SelectNodes('./w:r', $ns))) {
        if ($r.SelectSingleNode('./w:t', $ns)) { $textRun = $r; break }
    }
    if ($null -ne $textRun) {
        $p.InsertBefore($frag, $textRun) | Out-Null
    }
    else {
        $p.AppendChild($frag) | Out-Null
    }

    if (-not [string]::IsNullOrWhiteSpace($NewText)) {
        Set-ParaText -p $p -Text $NewText | Out-Null
    }
    Write-Output ("OK рисунок в абзаце '" + $ParaMatch + "' -> " + $media.Media)
}

# Вставляет абзац-подпись сразу после абзаца-маркера.
function Insert-CaptionAfter {
    param(
        $Ctx,
        [string]$AfterMatch,
        [string]$CaptionText,
        [string]$TemplateCaptionMatch,
        [int]$Occurrence = 1
    )
    $anchor = Find-Para -Ctx $Ctx -Match $AfterMatch -Occurrence $Occurrence -BodyLevel
    if ($null -eq $anchor) { throw ("Абзац-маркер не найден: '" + $AfterMatch + "'") }
    $tpl = Find-Para -Ctx $Ctx -Match $TemplateCaptionMatch -BodyLevel
    if ($null -eq $tpl) { throw ("Шаблон подписи не найден: '" + $TemplateCaptionMatch + "'") }
    $cap = $tpl.CloneNode($true)
    $ns = New-Object System.Xml.XmlNamespaceManager($Ctx.Doc.NameTable)
    $ns.AddNamespace('w', $script:W)
    foreach ($d in @($cap.SelectNodes('.//w:drawing', $ns))) { $d.ParentNode.RemoveChild($d) | Out-Null }
    foreach ($o in @($cap.SelectNodes('.//w:object', $ns))) { $o.ParentNode.RemoveChild($o) | Out-Null }
    Set-ParaText -p $cap -Text $CaptionText | Out-Null
    $anchor.ParentNode.InsertAfter($cap, $anchor) | Out-Null
    Write-Output ('OK вставлена подпись: ' + $CaptionText)
}

# Приводит абзац к виду стилевого заголовка (нужно для оглавления).
function Set-HeadingStyle {
    param($Ctx, [string]$Match, [string]$StyleId = '2')

    $p = Find-Para -Ctx $Ctx -Match $Match -BodyLevel
    if ($null -eq $p) { throw ("Абзац не найден: '" + $Match + "'") }

    $ns = New-Object System.Xml.XmlNamespaceManager($p.OwnerDocument.NameTable)
    $ns.AddNamespace('w', $script:W)

    $pPr = $p.SelectSingleNode('./w:pPr', $ns)
    if ($null -eq $pPr) {
        $pPr = $p.OwnerDocument.CreateElement('w', 'pPr', $script:W)
        if ($null -ne $p.FirstChild) { $p.InsertBefore($pPr, $p.FirstChild) | Out-Null }
        else { $p.AppendChild($pPr) | Out-Null }
    }
    while ($pPr.HasChildNodes) { $pPr.RemoveChild($pPr.FirstChild) | Out-Null }

    $st = $p.OwnerDocument.CreateElement('w', 'pStyle', $script:W)
    $a = $p.OwnerDocument.CreateAttribute('w', 'val', $script:W); $a.Value = $StyleId
    $st.Attributes.Append($a) | Out-Null
    $pPr.AppendChild($st) | Out-Null

    $pPr.AppendChild($p.OwnerDocument.CreateElement('w', 'keepNext', $script:W)) | Out-Null

    $sp = $p.OwnerDocument.CreateElement('w', 'spacing', $script:W)
    $a2 = $p.OwnerDocument.CreateAttribute('w', 'after', $script:W); $a2.Value = '120'
    $sp.Attributes.Append($a2) | Out-Null
    $pPr.AppendChild($sp) | Out-Null

    # убираем ручное жирное начертание — его даёт стиль
    foreach ($b in @($p.SelectNodes('.//w:rPr/w:b', $ns))) { $b.ParentNode.RemoveChild($b) | Out-Null }

    Write-Output ("OK стиль заголовка '" + $StyleId + "' применён: " + $Match)
}

# Удаляет абзац, стоящий непосредственно перед абзацем-маркером.
function Remove-ParaBefore {
    param($Ctx, [string]$AnchorMatch)
    $anchor = Find-Para -Ctx $Ctx -Match $AnchorMatch -BodyLevel
    if ($null -eq $anchor) { throw ("Абзац-маркер не найден: '" + $AnchorMatch + "'") }
    $prev = $anchor.PreviousSibling
    while ($null -ne $prev -and -not ($prev.NodeType -eq 'Element' -and $prev.LocalName -eq 'p')) {
        $prev = $prev.PreviousSibling
    }
    if ($null -eq $prev) { throw ('Перед абзацем нет другого абзаца') }
    $prev.ParentNode.RemoveChild($prev) | Out-Null
    Write-Output ("OK удалён абзац перед: " + $AnchorMatch)
}

# Меняет текст абзаца, стоящего непосредственно перед абзацем-маркером.
function Set-ParaBefore {
    param($Ctx, [string]$AnchorMatch, [string]$NewText)
    $anchor = Find-Para -Ctx $Ctx -Match $AnchorMatch -BodyLevel
    if ($null -eq $anchor) { throw ("Абзац-маркер не найден: '" + $AnchorMatch + "'") }
    $prev = $anchor.PreviousSibling
    while ($null -ne $prev -and -not ($prev.NodeType -eq 'Element' -and $prev.LocalName -eq 'p')) {
        $prev = $prev.PreviousSibling
    }
    if ($null -eq $prev) { throw ('Перед абзацем нет другого абзаца') }
    Set-ParaText -p $prev -Text $NewText | Out-Null
    Write-Output ('OK изменён абзац перед: ' + $AnchorMatch)
}
