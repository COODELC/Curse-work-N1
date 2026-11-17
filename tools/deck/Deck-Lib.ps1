# -*- coding: utf-8 -*-
# Библиотека сборки презентации через PowerPoint COM.
# Слайд 16:9 = 960 x 540 pt.

Add-Type -AssemblyName System.Drawing

$script:SW = 960
$script:SH = 540

function New-Rgb {
    param([int]$R, [int]$G, [int]$B)
    return ($R + ($G * 256) + ($B * 65536))
}

$script:CLR = @{
    Primary  = (New-Rgb 31 56 100)    # тёмно-синий — заголовки
    Accent   = (New-Rgb 46 116 181)   # синий — акценты
    Orange   = (New-Rgb 197 90 17)    # оранжевый — акцент 2
    Green    = (New-Rgb 56 118 29)
    Light    = (New-Rgb 242 246 251)  # подложка карточек
    Line     = (New-Rgb 200 212 228)
    Gray     = (New-Rgb 89 89 89)
    White    = (New-Rgb 255 255 255)
    Dark     = (New-Rgb 38 38 38)
    CodeBg   = (New-Rgb 45 45 48)
}

$script:FontTitle = 'Segoe UI Semibold'
$script:FontBody  = 'Segoe UI'
$script:FontCode  = 'Consolas'

function New-Deck {
    param()
    $ppt = New-Object -ComObject PowerPoint.Application
    $ppt.Visible = -1
    $pres = $ppt.Presentations.Add(-1)
    # 16:9 — ширина/высота в пунктах
    $pres.PageSetup.SlideSize = 15   # ppSlideSizeOnScreen16x9
    $pres.PageSetup.SlideWidth = $script:SW
    $pres.PageSetup.SlideHeight = $script:SH
    return [pscustomobject]@{ App = $ppt; Pres = $pres; Map = (New-Object System.Collections.ArrayList) }
}

# Создаёт пустой слайд и возвращает его.
function Add-BlankSlide {
    param($Deck)
    return $Deck.Pres.Slides.Add($Deck.Pres.Slides.Count + 1, 12)  # ppLayoutBlank
}

# Заголовок + подзаголовок + линейка. Возвращает y низа шапки.
function Add-SlideHeader {
    param($Slide, [string]$Title, [string]$Subtitle = '', [switch]$NoLine)
    $sh = $Slide.Shapes
    # акцентная полоса слева
    $bar = $sh.AddShape(1, 44, 34, 7, 40)   # msoShapeRectangle
    $bar.Fill.ForeColor.RGB = $script:CLR.Accent
    $bar.Line.Visible = 0
    $bar.Shadow.Visible = 0

    $tb = $sh.AddTextbox(1, 62, 28, 860, 46)
    $tr = $tb.TextFrame.TextRange
    $tr.Text = $Title
    $tr.Font.Name = $script:FontTitle
    $tr.Font.Size = 26
    $tr.Font.Bold = -1
    $tr.Font.Color.RGB = $script:CLR.Primary
    $tb.TextFrame.WordWrap = -1
    $tb.TextFrame.AutoSize = 0

    $y = 82
    if ($Subtitle -ne '') {
        $sb = $sh.AddTextbox(1, 64, 72, 850, 24)
        $sr = $sb.TextFrame.TextRange
        $sr.Text = $Subtitle
        $sr.Font.Name = $script:FontBody
        $sr.Font.Size = 13
        $sr.Font.Color.RGB = $script:CLR.Gray
        $sb.TextFrame.WordWrap = -1
        $y = 102
    }
    if (-not $NoLine) {
        $ln = $sh.AddShape(1, 44, $y, 872, 1.4)
        $ln.Fill.ForeColor.RGB = $script:CLR.Line
        $ln.Line.Visible = 0
        $ln.Shadow.Visible = 0
        $y = $y + 14
    }
    return $y
}

# Список маркеров. $items — массив строк или @{Text=..;Level=..;Bold=$true}
function Add-Bullets {
    param(
        $Slide, [object[]]$Items,
        [double]$Left = 50, [double]$Top = 110, [double]$Width = 860, [double]$Height = 380,
        [double]$FontSize = 16, [double]$LineSpacing = 1.0
    )
    $tb = $Slide.Shapes.AddTextbox(1, $Left, $Top, $Width, $Height)
    $tf = $tb.TextFrame
    $tf.WordWrap = -1
    $tr = $tf.TextRange

    # сначала весь текст — в PowerPoint абзацы добавляются только так
    $texts = New-Object System.Collections.ArrayList
    $meta = New-Object System.Collections.ArrayList
    foreach ($it in $Items) {
        $txt = ''
        $lvl = 0
        $bold = $false
        $col = $null
        if ($it -is [string]) { $txt = $it }
        else {
            $txt = $it.Text
            if ($it.ContainsKey('Level')) { $lvl = [int]$it.Level }
            if ($it.ContainsKey('Bold')) { $bold = [bool]$it.Bold }
            if ($it.ContainsKey('Color')) { $col = $it.Color }
        }
        [void]$texts.Add($txt)
        [void]$meta.Add([pscustomobject]@{ Level = $lvl; Bold = $bold; Color = $col })
    }
    $tr.Text = ($texts -join "`r")

    # абзацы берём через TextFrame2 — только там Paragraphs() отдаёт коллекцию
    $paras = $tb.TextFrame2.TextRange.Paragraphs()
    $count = $paras.Count
    for ($i = 1; $i -le $count; $i++) {
        $m = $meta[$i - 1]
        $p = $paras.Item($i)
        try { $p.ParagraphFormat.IndentLevel = $m.Level + 1 } catch { }
        $p.Font.Name = $script:FontBody
        $p.Font.Size = $FontSize - ($m.Level * 2)
        $p.Font.Bold = $(if ($m.Bold) { -1 } else { 0 })
        $p.Font.Fill.ForeColor.RGB = $(if ($null -ne $m.Color) { $m.Color } else { $script:CLR.Dark })
        $p.ParagraphFormat.SpaceAfter = 8
        $p.ParagraphFormat.SpaceBefore = 0
        # ВАЖНО: при LineRuleWithin = true SpaceWithin задаётся МНОЖИТЕЛЕМ
        # высоты строки, а не в пунктах (17.5 = 17,5 межстрочных интервалов).
        $p.ParagraphFormat.LineRuleWithin = -1
        $p.ParagraphFormat.SpaceWithin = 1.15 * $LineSpacing
        $p.ParagraphFormat.Bullet.Visible = -1
        $p.ParagraphFormat.Bullet.Type = 1
        $p.ParagraphFormat.Bullet.Character = 8226
        $p.ParagraphFormat.Bullet.Font.Fill.ForeColor.RGB = $script:CLR.Accent
    }
    # отступы уровней задаются через Ruler текстового фрейма
    for ($lv = 1; $lv -le 3; $lv++) {
        try {
            $lvlRule = $tf.Ruler.Levels($lv)
            $lvlRule.FirstMargin = -18
            $lvlRule.LeftMargin = 18 + (($lv - 1) * 20)
        }
        catch { }
    }
    # фиксируем размер рамки — иначе PowerPoint растянет её по содержимому
    $tf.AutoSize = 0
    $tb.Height = $Height
    return $tb
}

# Карточка-подложка.
function Add-Card {
    param($Slide, [double]$Left, [double]$Top, [double]$Width, [double]$Height, [int]$Fill = -1)
    $c = $Slide.Shapes.AddShape(5, $Left, $Top, $Width, $Height)  # roundedRect
    $c.Fill.ForeColor.RGB = $(if ($Fill -ge 0) { $Fill } else { $script:CLR.Light })
    $c.Line.ForeColor.RGB = $script:CLR.Line
    $c.Line.Weight = 0.75
    $c.Shadow.Visible = 0
    $c.Adjustments.Item(1) = 0.06
    return $c
}

# Изображение с вписыванием в область по пропорциям. Возвращает фигуру.
function Add-PictureFit {
    param(
        $Slide, [string]$Path,
        [double]$Left, [double]$Top, [double]$MaxWidth, [double]$MaxHeight
    )
    $img = [System.Drawing.Image]::FromFile($Path)
    try {
        $ar = $img.Width / $img.Height
    }
    finally { $img.Dispose() }
    $w = $MaxWidth
    $h = $w / $ar
    if ($h -gt $MaxHeight) { $h = $MaxHeight; $w = $h * $ar }
    $x = $Left + (($MaxWidth - $w) / 2)
    $y = $Top + (($MaxHeight - $h) / 2)
    $shp = $Slide.Shapes.AddPicture($Path, 0, -1, $x, $y, $w, $h)
    $shp.Line.Visible = -1
    $shp.Line.ForeColor.RGB = $script:CLR.Line
    $shp.Line.Weight = 0.75
    return $shp
}

# Таблица. $Rows — массив массивов строк; $Head — первая строка как заголовок.
function Add-Table {
    param(
        $Slide, [object[]]$Rows,
        [double]$Left, [double]$Top, [double]$Width, [double]$Height,
        [double]$FontSize = 13, [double[]]$ColWidths = $null,
        [switch]$NoHeader
    )
    $r = $Rows.Count
    $c = $Rows[0].Count
    $tblShape = $Slide.Shapes.AddTable($r, $c, $Left, $Top, $Width, $Height)
    $tbl = $tblShape.Table
    if ($null -ne $ColWidths) {
        for ($i = 0; $i -lt $c -and $i -lt $ColWidths.Count; $i++) {
            $tbl.Columns($i + 1).Width = $ColWidths[$i]
        }
    }
    for ($ri = 0; $ri -lt $r; $ri++) {
        $row = $Rows[$ri]
        $isHead = ($ri -eq 0 -and -not $NoHeader)
        for ($ci = 0; $ci -lt $c; $ci++) {
            $val = ''
            if ($ci -lt $row.Count) { $val = [string]$row[$ci] }
            $cell = $tbl.Cell($ri + 1, $ci + 1)
            $cell.Shape.TextFrame.TextRange.Text = $val
            $tr = $cell.Shape.TextFrame.TextRange
            $tr.Font.Name = $script:FontBody
            $tr.Font.Size = $FontSize
            if ($isHead) {
                $tr.Font.Bold = -1
                $tr.Font.Color.RGB = $script:CLR.White
                $cell.Shape.Fill.ForeColor.RGB = $script:CLR.Primary
            }
            else {
                $tr.Font.Bold = 0
                $tr.Font.Color.RGB = $script:CLR.Dark
                if (($ri % 2) -eq 0) {
                    $cell.Shape.Fill.ForeColor.RGB = $script:CLR.Light
                }
                else {
                    $cell.Shape.Fill.ForeColor.RGB = $script:CLR.White
                }
            }
            $cell.Shape.TextFrame.MarginLeft = 6
            $cell.Shape.TextFrame.MarginRight = 6
            $cell.Shape.TextFrame.MarginTop = 3
            $cell.Shape.TextFrame.MarginBottom = 3
            $cell.Shape.TextFrame.VerticalAnchor = 3  # msoAnchorMiddle
        }
        $tbl.Rows($ri + 1).Height = $(if ($isHead) { 26 } else { 22 })
    }
    return $tbl
}

# Подпись «Рисунок N — ...» под картинкой.
function Add-Caption {
    param($Slide, [string]$Text, [double]$Left, [double]$Top, [double]$Width = 860, [double]$Size = 11)
    $tb = $Slide.Shapes.AddTextbox(1, $Left, $Top, $Width, 20)
    $tr = $tb.TextFrame.TextRange
    $tr.Text = $Text
    $tr.Font.Name = $script:FontBody
    $tr.Font.Size = $Size
    $tr.Font.Italic = -1
    $tr.Font.Color.RGB = $script:CLR.Gray
    $tb.TextFrame.WordWrap = -1
    return $tb
}

# Номер страницы в правом нижнем углу.
function Add-PageNumber {
    param($Slide, [int]$Number, [string]$Total = '')
    $tb = $Slide.Shapes.AddTextbox(1, 866, 505, 52, 20)
    $tr = $tb.TextFrame.TextRange
    $tr.Text = "$Number"
    $tr.Font.Name = $script:FontBody
    $tr.Font.Size = 11
    $tr.Font.Color.RGB = $script:CLR.Gray
    $tb.TextFrame.TextRange.ParagraphFormat.Alignment = 3
    if ($Total -ne '') {
        $tb2 = $Slide.Shapes.AddTextbox(1, 830, 505, 32, 20)
        $t2 = $tb2.TextFrame.TextRange
        $t2.Text = "/ $Total"
        $t2.Font.Name = $script:FontBody
        $t2.Font.Size = 11
        $t2.Font.Color.RGB = $script:CLR.Line
        $tb2.TextFrame.TextRange.ParagraphFormat.Alignment = 2
    }
}

# Нижний колонтитул с названием работы.
function Add-Footer {
    param($Slide, [string]$Text = 'Информационная система для магазина автозапчастей')
    $tb = $Slide.Shapes.AddTextbox(1, 44, 505, 500, 18)
    $tr = $tb.TextFrame.TextRange
    $tr.Text = $Text
    $tr.Font.Name = $script:FontBody
    $tr.Font.Size = 10
    $tr.Font.Color.RGB = $script:CLR.Line
}

# Заметки докладчика.
function Add-SpeakerNote {
    param($Slide, [string]$Text)
    $Slide.NotesPage.Shapes.Item(2).TextFrame.TextRange.Text = $Text
}

# Регистрация элемента в карте.
function Add-MapEntry {
    param($Deck, [hashtable]$Entry)
    [void]$Deck.Map.Add($Entry)
}
