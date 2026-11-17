# -*- coding: utf-8 -*-
# Библиотека сборки презентации к защите через PowerPoint COM.
#
# Цветовая схема взята из пояснительной записки (таблица 15 «Цветовая схема
# приложения»): индиго #4F46E5, фиолетовый #7C3AED, розовый #DB2777,
# светлый фон #F8FAFC, основной текст #111827, вторичный #64748B,
# границы #E2E8F0, успех #059669, ошибка #E11D48.
#
# Слайд 16:9 = 960 x 540 pt.

Add-Type -AssemblyName System.Drawing

$script:SW = 960
$script:SH = 540

function New-Rgb {
    param([int]$R, [int]$G, [int]$B)
    return ($R + ($G * 256) + ($B * 65536))
}

$script:CLR = @{
    Primary   = (New-Rgb 79 70 229)     # #4F46E5 — индиго, акцентный
    Secondary = (New-Rgb 124 58 237)    # #7C3AED — фиолетовый
    Accent    = (New-Rgb 219 39 119)    # #DB2777 — розовый
    Success   = (New-Rgb 5 150 105)     # #059669
    Error     = (New-Rgb 225 29 72)     # #E11D48
    Info      = (New-Rgb 2 132 199)     # #0284C7
    Bg        = (New-Rgb 248 250 252)   # #F8FAFC
    BgSoft    = (New-Rgb 238 242 255)   # #EEF2FF
    Card      = (New-Rgb 255 255 255)   # #FFFFFF
    Line      = (New-Rgb 226 232 240)   # #E2E8F0
    Text      = (New-Rgb 17 24 39)      # #111827
    Muted     = (New-Rgb 100 116 139)   # #64748B
    CodeBg    = (New-Rgb 30 27 56)      # тёмный фон блока кода
    White     = (New-Rgb 255 255 255)
}

$script:FontTitle = 'Segoe UI Semibold'
$script:FontBody  = 'Segoe UI'
$script:FontCode  = 'Consolas'

function New-Deck {
    param()
    $ppt = New-Object -ComObject PowerPoint.Application
    $ppt.Visible = -1
    $pres = $ppt.Presentations.Add(-1)
    $pres.PageSetup.SlideSize = 15   # ppSlideSizeOnScreen16x9
    $pres.PageSetup.SlideWidth = $script:SW
    $pres.PageSetup.SlideHeight = $script:SH
    return [pscustomobject]@{ App = $ppt; Pres = $pres }
}

function Add-BlankSlide {
    param($Deck)
    return $Deck.Pres.Slides.Add($Deck.Pres.Slides.Count + 1, 12)  # ppLayoutBlank
}

# Шапка слайда: акцентная полоса, заголовок, подзаголовок, линейка.
function Add-SlideHeader {
    param($Slide, [string]$Title, [string]$Subtitle = '')

    $bar = $Slide.Shapes.AddShape(1, 44, 30, 6, 42)
    $bar.Fill.ForeColor.RGB = $script:CLR.Accent
    $bar.Line.Visible = 0
    $bar.Shadow.Visible = 0

    $tb = $Slide.Shapes.AddTextbox(1, 60, 24, 856, 48)
    $tr = $tb.TextFrame.TextRange
    $tr.Text = $Title
    $tr.Font.Name = $script:FontTitle
    $tr.Font.Size = 25
    $tr.Font.Bold = -1
    $tr.Font.Color.RGB = $script:CLR.Text
    $tb.TextFrame.WordWrap = -1
    $tb.TextFrame.AutoSize = 0

    $y = 80
    if ($Subtitle -ne '') {
        $sb = $Slide.Shapes.AddTextbox(1, 62, 66, 850, 24)
        $sr = $sb.TextFrame.TextRange
        $sr.Text = $Subtitle
        $sr.Font.Name = $script:FontBody
        $sr.Font.Size = 13
        $sr.Font.Color.RGB = $script:CLR.Muted
        $sb.TextFrame.WordWrap = -1
        $y = 94
    }

    $ln = $Slide.Shapes.AddShape(1, 44, $y, 872, 1.4)
    $ln.Fill.ForeColor.RGB = $script:CLR.Line
    $ln.Line.Visible = 0
    $ln.Shadow.Visible = 0

    return ($y + 14)
}

# Текстовый блок из абзацев. $items — строки или @{Text=..;Bold=..;Color=..;Level=..}
function Add-Paragraphs {
    param(
        $Slide, [object[]]$Items,
        [double]$Left = 50, [double]$Top = 110, [double]$Width = 860, [double]$Height = 380,
        [double]$FontSize = 16, [double]$Spacing = 1.0, [bool]$Bullets = $true
    )
    $tb = $Slide.Shapes.AddTextbox(1, $Left, $Top, $Width, $Height)
    $tf = $tb.TextFrame
    $tf.WordWrap = -1
    $tr = $tf.TextRange

    $texts = New-Object System.Collections.ArrayList
    $meta = New-Object System.Collections.ArrayList

    foreach ($it in $Items) {
        $txt = ''; $lvl = 0; $bold = $false; $col = $null
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

    $paras = $tb.TextFrame2.TextRange.Paragraphs()
    for ($i = 1; $i -le $paras.Count; $i++) {
        $m = $meta[$i - 1]
        $p = $paras.Item($i)
        try { $p.ParagraphFormat.IndentLevel = $m.Level + 1 } catch { }
        $p.Font.Name = $script:FontBody
        $p.Font.Size = $FontSize - ($m.Level * 2)
        $p.Font.Bold = $(if ($m.Bold) { -1 } else { 0 })
        $p.Font.Fill.ForeColor.RGB = $(if ($null -ne $m.Color) { $m.Color } else { $script:CLR.Text })
        $p.ParagraphFormat.SpaceAfter = 7
        $p.ParagraphFormat.SpaceBefore = 0
        $p.ParagraphFormat.LineRuleWithin = -1
        $p.ParagraphFormat.SpaceWithin = 1.12 * $Spacing
        if ($Bullets) {
            $p.ParagraphFormat.Bullet.Visible = -1
            $p.ParagraphFormat.Bullet.Type = 1
            $p.ParagraphFormat.Bullet.Character = 8226
            $p.ParagraphFormat.Bullet.Font.Fill.ForeColor.RGB = $script:CLR.Accent
        }
        else {
            $p.ParagraphFormat.Bullet.Visible = 0
        }
    }

    for ($lv = 1; $lv -le 3; $lv++) {
        try {
            $lvlRule = $tf.Ruler.Levels($lv)
            $lvlRule.FirstMargin = -18
            $lvlRule.LeftMargin = 18 + (($lv - 1) * 20)
        }
        catch { }
    }

    $tf.AutoSize = 0
    $tb.Height = $Height
    return $tb
}

# Карточка-подложка со скруглением.
function Add-Card {
    param($Slide, [double]$Left, [double]$Top, [double]$Width, [double]$Height, [int]$Fill = -1)
    $c = $Slide.Shapes.AddShape(5, $Left, $Top, $Width, $Height)  # roundedRect
    $c.Fill.ForeColor.RGB = $(if ($Fill -ge 0) { $Fill } else { $script:CLR.Card })
    $c.Line.ForeColor.RGB = $script:CLR.Line
    $c.Line.Weight = 0.75
    $c.Shadow.Visible = 0
    try { $c.Adjustments.Item(1) = 0.06 } catch { }
    return $c
}

# Изображение, вписанное в прямоугольник по пропорциям. Возвращает фигуру.
function Add-PictureFit {
    param(
        $Slide, [string]$Path,
        [double]$Left, [double]$Top, [double]$MaxWidth, [double]$MaxHeight,
        [switch]$NoBorder
    )
    $img = [System.Drawing.Image]::FromFile($Path)
    try { $ar = $img.Width / $img.Height }
    finally { $img.Dispose() }

    $w = $MaxWidth
    $h = $w / $ar
    if ($h -gt $MaxHeight) { $h = $MaxHeight; $w = $h * $ar }

    $x = $Left + (($MaxWidth - $w) / 2)
    $y = $Top + (($MaxHeight - $h) / 2)

    $shp = $Slide.Shapes.AddPicture($Path, 0, -1, $x, $y, $w, $h)

    if (-not $NoBorder) {
        $shp.Line.Visible = -1
        $shp.Line.ForeColor.RGB = $script:CLR.Line
        $shp.Line.Weight = 0.75
    }

    return [pscustomobject]@{ Shape = $shp; X = $x; Y = $y; W = $w; H = $h }
}

# Выноска на снимке: рамка вокруг элемента + номерной кружок.
# $Rect — координаты элемента в пикселях исходного PNG.
# $Pic  — результат Add-PictureFit (положение и размер снимка на слайде).
#
# $Anchor задаёт положение номерного кружка:
#   'tr','tl','br','bl' — кружок наезжает на соответствующий угол рамки
#                         (соседние элементы снимка не перекрываются);
#   'right','left'      — кружок сбоку от рамки по вертикальному центру.
function Add-Callout {
    param(
        $Slide, $Pic, [string]$ImagePath, $Rect,
        [int]$Number, [string]$Label = '', [string]$Side = 'right', [string]$Anchor = 'tr'
    )
    if ($null -eq $Rect) { return $null }

    $src = [System.Drawing.Image]::FromFile($ImagePath)
    try { $srcW = $src.Width; $srcH = $src.Height }
    finally { $src.Dispose() }

    $scale = $Pic.W / $srcW

    $x = $Pic.X + ($Rect.x * $scale)
    $y = $Pic.Y + ($Rect.y * $scale)
    $w = $Rect.w * $scale
    $h = $Rect.h * $scale

    # рамка вокруг элемента
    $frame = $Slide.Shapes.AddShape(5, $x, $y, $w, $h)   # roundedRect
    $frame.Fill.Visible = 0
    $frame.Line.Visible = -1
    $frame.Line.ForeColor.RGB = $script:CLR.Accent
    $frame.Line.Weight = 2.25
    $frame.Shadow.Visible = 0
    try { $frame.Adjustments.Item(1) = 0.12 } catch { }

    # номерной кружок
    $d = 17.0
    $half = $d / 2

    $corners = @(
        @{ a = 'tr'; x = $x + $w - $half; y = $y - $half },
        @{ a = 'tl'; x = $x - $half;     y = $y - $half },
        @{ a = 'br'; x = $x + $w - $half; y = $y + $h - $half },
        @{ a = 'bl'; x = $x - $half;     y = $y + $h - $half }
    )

    $cx = 0.0
    $cy = 0.0

    if ($Anchor -eq 'right' -or $Anchor -eq 'left') {
        if ($Anchor -eq 'right') { $cx = $x + $w + 9 } else { $cx = $x - $d - 9 }
        $cy = $y + ($h / 2) - $half

        # если кружок вышел за границы слайда — переносим на другую сторону рамки
        if (($cx + $d) -gt ($script:SW - 8)) { $cx = $x - $d - 9 }
        if ($cx -lt 8) { $cx = $x + $w + 9 }
    }
    else {
        # сначала пробуем заданный угол, затем остальные — по порядку,
        # чтобы кружок не выходил за границы слайда
        $ordered = @($corners | Where-Object { $_.a -eq $Anchor }) + @($corners | Where-Object { $_.a -ne $Anchor })

        foreach ($c in $ordered) {
            $tx = [double]$c.x
            $ty = [double]$c.y

            if ($tx -ge 6 -and ($tx + $d) -le ($script:SW - 6) -and $ty -ge 6 -and ($ty + $d) -le ($script:SH - 6)) {
                $cx = $tx
                $cy = $ty
                break
            }
        }

        if ($cx -eq 0.0) {
            $first = $ordered[0]
            $cx = [double]$first.x
            $cy = [double]$first.y
        }
    }

    # круг не должен выходить за границы слайда
    if ($cx -lt 8) { $cx = 8 }
    if (($cx + $d) -gt ($script:SW - 8)) { $cx = $script:SW - 8 - $d }
    if ($cy -lt 8) { $cy = 8 }
    if (($cy + $d) -gt ($script:SH - 8)) { $cy = $script:SH - 8 - $d }

    $dot = $Slide.Shapes.AddShape(9, $cx, $cy, $d, $d)   # msoShapeOval
    $dot.Fill.ForeColor.RGB = $script:CLR.Accent
    $dot.Line.Visible = -1
    $dot.Line.ForeColor.RGB = $script:CLR.White
    $dot.Line.Weight = 1.25
    $dot.Shadow.Visible = 0
    $dot.TextFrame.TextRange.Text = "$Number"
    $dot.TextFrame.TextRange.Font.Name = $script:FontBody
    $dot.TextFrame.TextRange.Font.Size = 11
    $dot.TextFrame.TextRange.Font.Bold = -1
    $dot.TextFrame.TextRange.Font.Color.RGB = $script:CLR.White
    $dot.TextFrame.TextRange.ParagraphFormat.Alignment = 2
    $dot.TextFrame.VerticalAnchor = 3
    $dot.TextFrame.MarginLeft = 0
    $dot.TextFrame.MarginRight = 0

    if ($Label -ne '') {
        $lw = 150.0
        $lx = if ($cx -ge $x) { $cx + $d + 6 } else { $cx - $lw - 6 }
        if ($lx + $lw -gt ($script:SW - 6)) { $lx = $script:SW - 6 - $lw }
        if ($lx -lt 6) { $lx = 6 }

        $lt = $Slide.Shapes.AddTextbox(1, $lx, ($cy - 4), $lw, 26)
        $ltr = $lt.TextFrame.TextRange
        $ltr.Text = "$Number. $Label"
        $ltr.Font.Name = $script:FontBody
        $ltr.Font.Size = 11
        $ltr.Font.Bold = -1
        $ltr.Font.Color.RGB = $script:CLR.Text
        $lt.TextFrame.WordWrap = -1
        $lt.TextFrame.AutoSize = 0
        $lt.Fill.ForeColor.RGB = $script:CLR.Card
        $lt.Fill.Transparency = 0.12
        $lt.Line.Visible = 0
    }

    return $frame
}

# Таблица. $Rows — массив массивов строк; первая строка — заголовок.
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
                $tr.Font.Color.RGB = $script:CLR.Text
                $cell.Shape.Fill.ForeColor.RGB = $(if (($ri % 2) -eq 0) { $script:CLR.BgSoft } else { $script:CLR.Card })
            }

            $cell.Shape.TextFrame.MarginLeft = 7
            $cell.Shape.TextFrame.MarginRight = 7
            $cell.Shape.TextFrame.MarginTop = 3
            $cell.Shape.TextFrame.MarginBottom = 3
            $cell.Shape.TextFrame.VerticalAnchor = 3
        }

        $tbl.Rows($ri + 1).Height = $(if ($isHead) { 25 } else { 21 })
    }

    return $tbl
}

# Блок программного кода на тёмном фоне. Кегль не меньше 14 пт.
function Add-CodeBlock {
    param(
        $Slide, [string[]]$Lines,
        [double]$Left, [double]$Top, [double]$Width, [double]$Height,
        [double]$FontSize = 14
    )
    $bg = $Slide.Shapes.AddShape(5, $Left, $Top, $Width, $Height)
    $bg.Fill.ForeColor.RGB = $script:CLR.CodeBg
    $bg.Line.Visible = 0
    $bg.Shadow.Visible = 0
    try { $bg.Adjustments.Item(1) = 0.04 } catch { }

    $tb = $Slide.Shapes.AddTextbox(1, ($Left + 14), ($Top + 10), ($Width - 28), ($Height - 20))
    $tf = $tb.TextFrame
    $tf.WordWrap = -1
    $tf.TextRange.Text = ($Lines -join "`r")

    $paras = $tb.TextFrame2.TextRange.Paragraphs()
    for ($i = 1; $i -le $paras.Count; $i++) {
        $p = $paras.Item($i)
        $p.Font.Name = $script:FontCode
        $p.Font.Size = $FontSize
        $p.Font.Fill.ForeColor.RGB = (New-Rgb 226 232 240)
        $p.ParagraphFormat.Bullet.Visible = 0
        $p.ParagraphFormat.SpaceAfter = 1
        $p.ParagraphFormat.SpaceBefore = 0
        $p.ParagraphFormat.LineRuleWithin = -1
        $p.ParagraphFormat.SpaceWithin = 1.0
    }

    $tf.AutoSize = 0
    $tb.Height = $Height - 20
    return $bg
}

# Подпись «Рисунок N — ...» под рисунком.
function Add-Caption {
    param($Slide, [string]$Text, [double]$Left, [double]$Top, [double]$Width = 860, [double]$Size = 11)
    $tb = $Slide.Shapes.AddTextbox(1, $Left, $Top, $Width, 20)
    $tr = $tb.TextFrame.TextRange
    $tr.Text = $Text
    $tr.Font.Name = $script:FontBody
    $tr.Font.Size = $Size
    $tr.Font.Italic = -1
    $tr.Font.Color.RGB = $script:CLR.Muted
    $tb.TextFrame.WordWrap = -1
    return $tb
}

# Крупная подпись-заголовок внутри карточки.
function Add-Label {
    param($Slide, [string]$Text, [double]$Left, [double]$Top, [double]$Width, [double]$Height,
        [double]$Size = 15, [bool]$Bold = $true, [int]$Color = -1)
    $tb = $Slide.Shapes.AddTextbox(1, $Left, $Top, $Width, $Height)
    $tr = $tb.TextFrame.TextRange
    $tr.Text = $Text
    $tr.Font.Name = $(if ($Bold) { $script:FontTitle } else { $script:FontBody })
    $tr.Font.Size = $Size
    $tr.Font.Bold = $(if ($Bold) { -1 } else { 0 })
    $tr.Font.Color.RGB = $(if ($Color -ge 0) { $Color } else { $script:CLR.Text })
    $tb.TextFrame.WordWrap = -1
    $tb.TextFrame.AutoSize = 0
    return $tb
}

# Текст карточки: заголовок и пояснение — в одном текстовом блоке,
# чтобы карточка считалась одним элементом текста на слайде.
function Add-CardText {
    param(
        $Slide, [string]$Title, [string]$Body,
        [double]$Left, [double]$Top, [double]$Width, [double]$Height,
        [double]$TitleSize = 13, [double]$BodySize = 12,
        [int]$TitleColor = -1, [int]$BodyColor = -1,
        [bool]$Center = $false
    )
    $tb = $Slide.Shapes.AddTextbox(1, $Left, $Top, $Width, $Height)
    $tf = $tb.TextFrame
    $tf.WordWrap = -1
    $tf.TextRange.Text = "$Title`r$Body"

    $paras = $tb.TextFrame2.TextRange.Paragraphs()

    if ($paras.Count -ge 1) {
        $p = $paras.Item(1)
        $p.Font.Name = $script:FontTitle
        $p.Font.Size = $TitleSize
        $p.Font.Bold = -1
        $p.Font.Fill.ForeColor.RGB = $(if ($TitleColor -ge 0) { $TitleColor } else { $script:CLR.Primary })
        $p.ParagraphFormat.SpaceAfter = 5
        $p.ParagraphFormat.LineRuleWithin = -1
        $p.ParagraphFormat.SpaceWithin = 1.05
        $p.ParagraphFormat.Bullet.Visible = 0
    }

    if ($paras.Count -ge 2) {
        $p = $paras.Item(2)
        $p.Font.Name = $script:FontBody
        $p.Font.Size = $BodySize
        $p.Font.Bold = 0
        $p.Font.Fill.ForeColor.RGB = $(if ($BodyColor -ge 0) { $BodyColor } else { $script:CLR.Text })
        $p.ParagraphFormat.SpaceAfter = 0
        $p.ParagraphFormat.LineRuleWithin = -1
        $p.ParagraphFormat.SpaceWithin = 1.08
        $p.ParagraphFormat.Bullet.Visible = 0
    }

    if ($Center) {
        $tf.TextRange.ParagraphFormat.Alignment = 2
    }

    $tf.AutoSize = 0
    $tb.Height = $Height
    return $tb
}

# Числовой показатель: крупное число + подпись.
function Add-Kpi {
    param($Slide, [string]$Value, [string]$Caption,
        [double]$Left, [double]$Top, [double]$Width, [double]$Height,
        [int]$Color = -1)
    $c = Add-Card -Slide $Slide -Left $Left -Top $Top -Width $Width -Height $Height
    $vb = Add-Label -Slide $Slide -Text $Value -Left ($Left + 12) -Top ($Top + 8) `
        -Width ($Width - 24) -Height 30 -Size 26 -Bold $true -Color $(if ($Color -ge 0) { $Color } else { $script:CLR.Primary })
    $vb.TextFrame.TextRange.ParagraphFormat.Alignment = 2
    $cb = Add-Label -Slide $Slide -Text $Caption -Left ($Left + 10) -Top ($Top + 40) `
        -Width ($Width - 20) -Height 24 -Size 12 -Bold $false -Color $script:CLR.Muted
    $cb.TextFrame.TextRange.ParagraphFormat.Alignment = 2
    return $c
}

function Add-PageNumber {
    param($Slide, [int]$Number, [string]$Total = '')
    $tb = $Slide.Shapes.AddTextbox(1, 866, 505, 52, 20)
    $tr = $tb.TextFrame.TextRange
    $tr.Text = "$Number"
    $tr.Font.Name = $script:FontBody
    $tr.Font.Size = 11
    $tr.Font.Color.RGB = $script:CLR.Muted
    $tr.ParagraphFormat.Alignment = 3

    if ($Total -ne '') {
        $tb2 = $Slide.Shapes.AddTextbox(1, 830, 505, 32, 20)
        $t2 = $tb2.TextFrame.TextRange
        $t2.Text = "/ $Total"
        $t2.Font.Name = $script:FontBody
        $t2.Font.Size = 11
        $t2.Font.Color.RGB = $script:CLR.Line
        $t2.ParagraphFormat.Alignment = 2
    }
}

function Add-Footer {
    param($Slide, [string]$Text = 'Информационная система для магазина автозапчастей')
    $tb = $Slide.Shapes.AddTextbox(1, 44, 505, 500, 18)
    $tr = $tb.TextFrame.TextRange
    $tr.Text = $Text
    $tr.Font.Name = $script:FontBody
    $tr.Font.Size = 10
    $tr.Font.Color.RGB = $script:CLR.Line
}

function Add-SpeakerNote {
    param($Slide, [string]$Text)
    $Slide.NotesPage.Shapes.Item(2).TextFrame.TextRange.Text = $Text
}
