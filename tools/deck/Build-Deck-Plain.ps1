# -*- coding: utf-8 -*-
# Текстовая презентация к защите курсового проекта — без оформления.
#
# Отличие от Build-Deck-Defense.ps1: здесь нет цветовой схемы, карточек,
# плашек и подписей к рисункам. Только заголовок, текст и изображения.
# Структура слайдов — по списку из методических указаний:
#   тема; проблема; цель и задачи; требования; модель данных;
#   архитектура; алгоритм заказа; каталог; оформление; роли; тесты; итоги.
#
# Запуск:
#   powershell -File tools/deck/Build-Deck-Plain.ps1

param(
    [string]$OutFile = 'Презентация_текстовая_Копейкин.pptx'
)

$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName System.Drawing

$root = Split-Path -Parent $PSScriptRoot
$root = Split-Path -Parent $root
$screens = Join-Path $root 'docs\screens'
$diag = Join-Path $root 'docs\diagrams'
$outPath = Join-Path $root $OutFile

$script:SW = 960
$script:SH = 540

$BLACK = 0
$WHITE = 16777215
$GRAY = 8421504

$script:FontText = 'Segoe UI'
$script:FontCode = 'Consolas'

# --- Заголовок слайда ---

function Add-Title {
    param($Slide, [string]$Text)

    $tb = $Slide.Shapes.AddTextbox(1, 50, 30, 860, 46)
    $tr = $tb.TextFrame.TextRange
    $tr.Text = $Text
    $tr.Font.Name = $script:FontText
    $tr.Font.Size = 26
    $tr.Font.Bold = -1
    $tr.Font.Color.RGB = $BLACK
    $tb.TextFrame.WordWrap = -1
    $tb.TextFrame.AutoSize = 0
    return $tb
}

# --- Текстовый блок: список пунктов ---

function Add-Bullets {
    param(
        $Slide, [string[]]$Items,
        [double]$Left = 50, [double]$Top = 100, [double]$Width = 860, [double]$Height = 380,
        [double]$FontSize = 16, [double]$Spacing = 9
    )

    $tb = $Slide.Shapes.AddTextbox(1, $Left, $Top, $Width, $Height)
    $tf = $tb.TextFrame
    $tf.WordWrap = -1

    $lines = New-Object System.Collections.ArrayList
    foreach ($item in $Items) { [void]$lines.Add("• $item") }

    $tf.TextRange.Text = ($lines -join "`r")

    $paras = $tb.TextFrame2.TextRange.Paragraphs()
    for ($i = 1; $i -le $paras.Count; $i++) {
        $p = $paras.Item($i)
        $p.Font.Name = $script:FontText
        $p.Font.Size = $FontSize
        $p.Font.Bold = 0
        $p.Font.Fill.ForeColor.RGB = $BLACK
        $p.ParagraphFormat.Bullet.Visible = 0
        $p.ParagraphFormat.SpaceAfter = $Spacing
        $p.ParagraphFormat.SpaceBefore = 0
        $p.ParagraphFormat.LineRuleWithin = -1
        $p.ParagraphFormat.SpaceWithin = 1.1
    }

    $tf.AutoSize = 0
    $tb.Height = $Height
    return $tb
}

# --- Один абзац текста (без маркера) ---

function Add-Text {
    param(
        $Slide, [string]$Text,
        [double]$Left, [double]$Top, [double]$Width, [double]$Height,
        [double]$FontSize = 16, [bool]$Bold = $false, [int]$Color = 0,
        [int]$Align = 1
    )

    $tb = $Slide.Shapes.AddTextbox(1, $Left, $Top, $Width, $Height)
    $tf = $tb.TextFrame
    $tf.WordWrap = -1
    $tf.TextRange.Text = $Text
    $tf.TextRange.Font.Name = $script:FontText
    $tf.TextRange.Font.Size = $FontSize
    $tf.TextRange.Font.Bold = $(if ($Bold) { -1 } else { 0 })
    $tf.TextRange.Font.Color.RGB = $Color
    $tf.TextRange.ParagraphFormat.Alignment = $Align
    $tf.AutoSize = 0
    $tb.Height = $Height
    return $tb
}

# --- Изображение без рамки и подписи ---

function Add-Picture {
    param(
        $Slide, [string]$Path,
        [double]$Left, [double]$Top, [double]$MaxWidth, [double]$MaxHeight
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

    return [pscustomobject]@{ Shape = $shp; X = $x; Y = $y; W = $w; H = $h }
}

# --- Выноска на снимке: рамка и номерной кружок ---

function Add-Callout {
    param(
        $Slide, $Pic, [string]$ImagePath, $Rect,
        [int]$Number, [string]$Anchor = 'tr'
    )
    if ($null -eq $Rect) { return $null }

    $src = [System.Drawing.Image]::FromFile($ImagePath)
    try { $srcW = $src.Width }
    finally { $src.Dispose() }

    $scale = $Pic.W / $srcW

    $x = $Pic.X + ($Rect.x * $scale)
    $y = $Pic.Y + ($Rect.y * $scale)
    $w = $Rect.w * $scale
    $h = $Rect.h * $scale

    $frame = $Slide.Shapes.AddShape(1, $x, $y, $w, $h)   # прямоугольник
    $frame.Fill.Visible = 0
    $frame.Line.Visible = -1
    $frame.Line.ForeColor.RGB = $BLACK
    $frame.Line.Weight = 1.25
    $frame.Shadow.Visible = 0

    $d = 15.0
    $half = $d / 2

    $corners = @(
        @{ a = 'tr'; x = $x + $w - $half; y = $y - $half },
        @{ a = 'tl'; x = $x - $half;     y = $y - $half },
        @{ a = 'br'; x = $x + $w - $half; y = $y + $h - $half },
        @{ a = 'bl'; x = $x - $half;     y = $y + $h - $half }
    )

    $cx = 0.0
    $cy = 0.0
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

    if ($cx -lt 8) { $cx = 8 }
    if (($cx + $d) -gt ($script:SW - 8)) { $cx = $script:SW - 8 - $d }
    if ($cy -lt 8) { $cy = 8 }
    if (($cy + $d) -gt ($script:SH - 8)) { $cy = $script:SH - 8 - $d }

    $dot = $Slide.Shapes.AddShape(9, $cx, $cy, $d, $d)
    $dot.Fill.ForeColor.RGB = $BLACK
    $dot.Line.Visible = 0
    $dot.Shadow.Visible = 0
    $dot.TextFrame.TextRange.Text = "$Number"
    $dot.TextFrame.TextRange.Font.Name = $script:FontText
    $dot.TextFrame.TextRange.Font.Size = 10
    $dot.TextFrame.TextRange.Font.Bold = -1
    $dot.TextFrame.TextRange.Font.Color.RGB = $WHITE
    $dot.TextFrame.TextRange.ParagraphFormat.Alignment = 2
    $dot.TextFrame.VerticalAnchor = 3
    $dot.TextFrame.MarginLeft = 0
    $dot.TextFrame.MarginRight = 0

    return $frame
}

# --- Блок кода: моноширинный текст, кегль не меньше 14 пт ---

function Add-Code {
    param(
        $Slide, [string[]]$Lines,
        [double]$Left, [double]$Top, [double]$Width, [double]$Height,
        [double]$FontSize = 14
    )

    $tb = $Slide.Shapes.AddTextbox(1, $Left, $Top, $Width, $Height)
    $tf = $tb.TextFrame
    $tf.WordWrap = 0
    $tf.TextRange.Text = ($Lines -join "`r")

    $paras = $tb.TextFrame2.TextRange.Paragraphs()
    for ($i = 1; $i -le $paras.Count; $i++) {
        $p = $paras.Item($i)
        $p.Font.Name = $script:FontCode
        $p.Font.Size = $FontSize
        $p.Font.Bold = 0
        $p.Font.Fill.ForeColor.RGB = $BLACK
        $p.ParagraphFormat.Bullet.Visible = 0
        $p.ParagraphFormat.SpaceAfter = 2
        $p.ParagraphFormat.SpaceBefore = 0
        $p.ParagraphFormat.LineRuleWithin = -1
        $p.ParagraphFormat.SpaceWithin = 1.0
    }

    $tf.AutoSize = 0
    $tb.Height = $Height
    return $tb
}

# --- Блок формул по ГОСТ: формула с номером и экспликацией ---

function Add-Formula {
    param(
        $Slide, [string]$Title, [string]$Formula, [string[]]$Legend,
        [double]$Left, [double]$Top, [double]$Width, [double]$Height
    )

    $tb = $Slide.Shapes.AddTextbox(1, $Left, $Top, $Width, $Height)
    $tf = $tb.TextFrame
    $tf.WordWrap = -1

    $lines = New-Object System.Collections.ArrayList
    [void]$lines.Add($Title)
    [void]$lines.Add($Formula)
    foreach ($l in $Legend) { [void]$lines.Add($l) }

    $tf.TextRange.Text = ($lines -join "`r")

    $paras = $tb.TextFrame2.TextRange.Paragraphs()
    for ($i = 1; $i -le $paras.Count; $i++) {
        $p = $paras.Item($i)
        $p.Font.Name = $script:FontText
        $p.Font.Size = $(if ($i -eq 2) { 15 } elseif ($i -eq 1) { 13 } else { 12 })
        $p.Font.Bold = $(if ($i -eq 2) { -1 } else { 0 })
        $p.Font.Fill.ForeColor.RGB = $BLACK
        $p.ParagraphFormat.Bullet.Visible = 0
        $p.ParagraphFormat.SpaceAfter = 3
        $p.ParagraphFormat.SpaceBefore = 0
        $p.ParagraphFormat.LineRuleWithin = -1
        $p.ParagraphFormat.SpaceWithin = 1.05
    }

    $tf.AutoSize = 0
    $tb.Height = $Height
    return $tb
}

function Add-PageNumber {
    param($Slide, [int]$Number)

    [void](Add-Text -Slide $Slide -Text "$Number" -Left 900 -Top 505 -Width 40 -Height 20 `
        -FontSize 11 -Bold $false -Color $GRAY -Align 3)
}

function Add-Note {
    param($Slide, [string]$Text)

    $Slide.NotesPage.Shapes.Item(2).TextFrame.TextRange.Text = $Text
}

# --- Подготовка ---

$callouts = @{}
$calloutFile = Join-Path $screens 'callouts.json'

if (Test-Path -LiteralPath $calloutFile) {
    $callouts = Get-Content -LiteralPath $calloutFile -Raw -Encoding UTF8 | ConvertFrom-Json
}

function Get-Rect {
    param([string]$Key)

    if (-not $callouts.PSObject.Properties.Name.Contains($Key)) { return $null }
    return $callouts.$Key
}

$notes = @{}

$notes['title'] = 'Здравствуйте! Меня зовут Андрей Копейкин, я студент группы 324. Тема моего курсового проекта — «Разработка информационной системы для магазина автозапчастей». Руководитель проекта — Т. Т. Дуларидзе.'
$notes['problem'] = 'Актуальность темы в том, что рынок автозапчастей очень большой: в каталоге одного магазина десятки и сотни позиций с разными артикулами, ценами и остатками. Когда учёт ведётся вручную или в таблицах, неизбежно возникают две проблемы: покупатель не получает точный остаток и сроки поставки, а продавец теряет заказы из-за ошибок и забытых записей. Кроме того, в магазине работают и продавцы, и менеджеры, и администратор, и обычный покупатель, — а значит, у каждого должен быть свой набор прав.'
$notes['goals'] = 'Цель работы — разработать информационную систему, которая автоматизирует учёт каталога автозапчастей, оформление заказов и работу персонала с разграничением прав доступа. Для достижения цели поставлены четыре задачи: спроектировать модель данных и физическую базу; разработать архитектуру приложения и алгоритм расчёта цены со скидкой; реализовать пользовательский интерфейс, включая корзину и оформление заказа; реализовать ролевую модель с четырьмя ролями и разграничением прав.'
$notes['requirements'] = 'Функциональные требования приведены в виде одиннадцати пунктов. Для неавторизованного пользователя: просмотр товаров, фильтрация по отделениям, выбор товаров, информация о магазине и кнопка входа. Для авторизованного: оформление заказа, отказ от заказа, отзывы и история покупок. Сотруднику доступна обработка заказов, администратору — управление базами данных магазина.'
$notes['model'] = 'В ходе анализа выделены четыре справочника — «Роли», «Категории», «Статусы заказа», — и четыре сущности: «Пользователи», «Товары», «Заказы» и «Отзывы». Всего семь таблиц, приведённых к третьей нормальной форме: нет транзитивных зависимостей и дублирования данных. Семь внешних ключей задают связи один-ко-многим и обеспечивают ссылочную целостность.'
$notes['physical'] = 'Физическая модель создаётся скриптом SQLCreateTable.sql: в нём описаны все семь таблиц с типами данных, внешними ключами и ограничениями NOT NULL. Данные загружаются скриптом insert_sqlserver.sql — всего 43 записи в справочниках, включая каталог из 12 позиций и 15 отзывов. Импорт выполняется одной командой, поэтому базу можно развернуть заново на любой машине.'
$notes['architecture'] = 'Приложение построено по четырёхуровневой архитектуре: интерфейс пользователя на Avalonia UI, логика окон на моделях представления, бизнес-правила в сервисах и доступ к данным через Entity Framework Core. Каждый уровень обращается только к следующему: представление никогда не обращается к базе напрямую. Благодаря этому сервисы тестируются без запуска интерфейса.'
$notes['algorithm'] = 'На схеме по ГОСТ 19.701-90 показан алгоритм оформления заказа. Сначала проверяется, что корзина не пуста, затем для каждой позиции — наличие товара и достаточное количество на складе. Если товара не хватает, выводится сообщение и цикл возвращается к выбору позиции. Когда все позиции обработаны, рассчитывается скидка и сумма, выполняется списание остатка, и только после этого сохраняются строки заказа и присваивается номер.'
$notes['code'] = 'Вот фрагмент метода CreateOrderAsync сервиса заказов. Сначала определяется, есть ли в истории покупателя заказы по данному товару за прошлый месяц: если их нет, применяется скидка 25 процентов, иначе цена остаётся базовой. Затем рассчитываются цена позиции и сумма заказа, после чего в рамках транзакции списывается остаток и добавляются строки заказа.'
$notes['catalog'] = 'Главная форма — это каталог автозапчастей. Поле поиска работает одновременно по названию и по артикулу, фильтр ограничивает каталог по категориям. Кнопка «Добавить в корзину» доступна у каждой карточки товара, кнопка «Корзина» ведёт к оформлению. В правом верхнем углу отображается текущий пользователь и его роль. Поиск и фильтр применяются одновременно, и выборка выполняется одним запросом к базе.'
$notes['order'] = 'Это корзина покупателя: кнопки «плюс» и «минус» меняют количество, «Удалить» возвращает товар в каталог. Кнопка «Оплатить» оформляет заказ и списывает остаток на складе, итоговая сумма учитывает скидку постоянного покупателя. Справа — история собственных заказов: пользователь видит только свои заказы, а отмена заказа возвращает товар на склад.'
$notes['roles'] = 'В системе четыре роли и девять операций. Гость только просматривает каталог. Покупатель дополнительно оформляет заказы, пишет отзывы, смотрит свою историю и меняет профиль. Менеджер дополнительно обрабатывает заказы. Администратор дополнительно управляет каталогом и пользователями. Роль определяется при входе и хранится в таблице пользователей, проверка выполняется методом Permissions.Allows и в интерфейсе, и в сервисе.'
$notes['tests'] = 'Проверка выполнена в двух видах: пятьдесят две модульные проверки сервисов и восемь сквозных проверок через UI Automation, которые повторяют реальный сценарий покупки. Ошибок нет. Отдельно проверены вход и разграничение прав, расчёт скидки двадцать пять процентов, списание и возврат остатка, а также работа всех четырёх ролей.'
$notes['result'] = 'Все поставленные задачи решены: проанализирована предметная область, спроектирована и наполнена база данных, реализованы архитектура и алгоритмы, разработан интерфейс и ролевая модель, продукт протестирован. Направления развития: печатные формы накладной и чека, резервное копирование базы и веб-версия приложения.'
$notes['end'] = 'Таким образом, цель курсового проекта достигнута, все задачи выполнены. Спасибо за внимание! Готов(а) ответить на ваши вопросы.'

# --- Сборка ---

$ppt = New-Object -ComObject PowerPoint.Application
$ppt.Visible = -1
$pres = $ppt.Presentations.Add(-1)
$pres.PageSetup.SlideSize = 15
$pres.PageSetup.SlideWidth = $script:SW
$pres.PageSetup.SlideHeight = $script:SH

$slides = New-Object System.Collections.ArrayList

function New-Slide {
    param([string]$Title, [string]$Note = '')

    $s = $pres.Slides.Add($pres.Slides.Count + 1, 12)   # ppLayoutBlank
    if ($Title -ne '') { [void](Add-Title -Slide $s -Text $Title) }
    if ($Note -ne '') { Add-Note -Slide $s -Text $Note }
    [void]$slides.Add($s)
    return $s
}

# ===== 1. Титульный слайд =====

$s = New-Slide -Note $notes['title']

[void](Add-Text -Slide $s -Text "Санкт-Петербургское государственное автономное`rпрофессиональное образовательное учреждение`r«Колледж туризма и прикладных технологий»" `
    -Left 90 -Top 46 -Width 780 -Height 70 -FontSize 15 -Bold $false -Color $BLACK -Align 2)
[void](Add-Text -Slide $s -Text 'Курсовой проект' -Left 90 -Top 136 -Width 780 -Height 28 -FontSize 16 -Bold $true -Color $BLACK -Align 2)
[void](Add-Text -Slide $s -Text 'на тему:' -Left 90 -Top 170 -Width 780 -Height 24 -FontSize 14 -Bold $false -Color $BLACK -Align 2)
[void](Add-Text -Slide $s -Text "«Разработка информационной системы`rдля магазина автозапчастей»" `
    -Left 90 -Top 196 -Width 780 -Height 64 -FontSize 23 -Bold $true -Color $BLACK -Align 2)
[void](Add-Text -Slide $s -Text 'МДК.01.01 Разработка программных модулей' `
    -Left 90 -Top 268 -Width 780 -Height 24 -FontSize 14 -Bold $false -Color $GRAY -Align 2)
[void](Add-Text -Slide $s -Text "Выполнил(а) студент(ка) гр. 324:`rА. А. Копейкин`rспециальность: 09.02.07`rИнформационные системы и программирование" `
    -Left 110 -Top 322 -Width 380 -Height 90 -FontSize 14 -Bold $false -Color $BLACK -Align 1)
[void](Add-Text -Slide $s -Text "Руководитель:`rТ. Т. Дуларидзе, преподаватель" `
    -Left 520 -Top 322 -Width 330 -Height 60 -FontSize 14 -Bold $false -Color $BLACK -Align 1)
[void](Add-Text -Slide $s -Text 'Санкт-Петербург, 2026' -Left 90 -Top 470 -Width 780 -Height 26 -FontSize 14 -Bold $false -Color $GRAY -Align 2)

# ===== 2. Тема и проблема =====

$s = New-Slide -Title 'Тема работы и постановка проблемы' -Note $notes['problem']

[void](Add-Bullets -Slide $s -Items @(
    'Тема: разработка информационной системы для магазина автозапчастей',
    'Проблема: ручной учёт каталога при большом ассортименте запчастей',
    'Покупатель не получает точный остаток и сроки поставки',
    'Продавец теряет заказы из-за ошибок и забытых записей',
    'У четырёх участников процесса должны быть разные права',
    'Объект: процесс продажи автозапчастей и учёт товарных остатков',
    'Предмет: программное обеспечение, автоматизирующее этот процесс'
) -Top 112)

# ===== 3. Цель и задачи =====

$s = New-Slide -Title 'Цель работы и поставленные задачи' -Note $notes['goals']

[void](Add-Bullets -Slide $s -Items @(
    'Цель: информационная система для учёта каталога, оформления заказов и работы персонала',
    'Задача 1: спроектировать модель данных, физическую базу и требования к системе',
    'Задача 2: разработать архитектуру приложения и алгоритм расчёта цены со скидкой',
    'Задача 3: реализовать интерфейс: каталог, корзина, оформление и отмена заказа',
    'Задача 4: реализовать ролевую модель: 4 роли и 9 операций',
    'Средства: C#, Avalonia UI, MVVM, Entity Framework Core, SQL Server, UI Automation'
) -Top 112)

# ===== 4. Функциональные требования =====

$s = New-Slide -Title 'Функциональные требования к системе' -Note $notes['requirements']

[void](Add-Bullets -Slide $s -Items @(
    'ФТ 1–5 (гость): просмотр товаров, фильтрация по отделениям, выбор товаров,',
    'информация о магазине, кнопка входа',
    'ФТ 6–9 (покупатель): оформление заказа, отказ от заказа, отзывы, история покупок',
    'ФТ 10 (сотрудник): обработка заказов — смена статуса и отмена с возвратом товара',
    'ФТ 11 (администратор): управление каталогом, категориями и пользователями',
    'Всего 11 функциональных требований, 4 роли пользователя'
) -Top 112)

# ===== 5. Модель данных =====

$s = New-Slide -Title 'Модель данных: ER-диаграмма' -Note $notes['model']

[void](Add-Bullets -Slide $s -Items @(
    '7 таблиц в третьей нормальной форме',
    'Справочники: Roles, Category, OrderStatus',
    'Сущности: Users, Item, Orders, ItemReview',
    '7 внешних ключей, связи один-ко-многим',
    'Аномалии вставки, удаления и обновления устранены',
    'Суррогатные ключи: Id_User, Id_Category, Id_Item, Id_Order'
) -Left 50 -Top 120 -Width 350 -Height 320 -FontSize 15)

[void](Add-Picture -Slide $s -Path (Join-Path $diag 'er.png') -Left 420 -Top 100 -MaxWidth 500 -MaxHeight 390)

# ===== 6. Физическая модель и импорт =====

$s = New-Slide -Title 'Физическая модель базы и результаты импорта' -Note $notes['physical']

[void](Add-Bullets -Slide $s -Items @(
    'Category — 6 категорий (отделения автозапчастей)',
    'Item — 12 позиций каталога: цена, остаток, описание, фото',
    'Users — 5 учётных записей, Roles — 4 роли',
    'Orders — 12 строк заказов со статусом и датой',
    'OrderStatus — 6 статусов, ItemReview — 15 отзывов',
    'Итого 43 записи в справочниках, база cursevork в SQL Server',
    'Создание — SQLCreateTable.sql, данные — insert_sqlserver.sql, импорт одной командой'
) -Top 112)

# ===== 7. Архитектура =====

$s = New-Slide -Title 'Архитектура приложения' -Note $notes['architecture']

[void](Add-Bullets -Slide $s -Items @(
    'Интерфейс пользователя — Avalonia UI, 9 экранных форм на XAML',
    'Логика окон — модели представления (ViewModel), шаблон MVVM',
    'Бизнес-правила — сервисы: заказы, каталог, отзывы, роли',
    'Доступ к данным — контекст Entity Framework Core, фабрика DbFactory',
    'Каждый уровень обращается только к следующему',
    'Сервисы тестируются без запуска интерфейса'
) -Left 50 -Top 120 -Width 420 -Height 320 -FontSize 15)

[void](Add-Picture -Slide $s -Path (Join-Path $diag 'architecture.png') -Left 500 -Top 100 -MaxWidth 420 -MaxHeight 390)

# ===== 8. Алгоритм оформления заказа =====

$s = New-Slide -Title 'Алгоритм оформления заказа' -Note $notes['algorithm']

[void](Add-Picture -Slide $s -Path (Join-Path $diag 'alg_create.png') -Left 50 -Top 100 -MaxWidth 470 -MaxHeight 380)

[void](Add-Bullets -Slide $s -Items @(
    'Проверка корзины, наличия товара и количества на складе',
    'Скидка 25 %, если за прошлый месяц по товару не было заказа',
    'Списание остатка и сохранение заказа — в одной транзакции',
    'Количество разворачивается в отдельные строки заказа',
    'Заказ свыше пяти позиций помечается как крупный',
    'Ошибка по остатку возвращает к выбору позиции'
) -Left 550 -Top 120 -Width 370 -Height 340 -FontSize 15)

# ===== 9. Формулы и код =====

$s = New-Slide -Title 'Расчёт цены: формулы и реализация' -Note $notes['code']

[void](Add-Formula -Slide $s -Title 'Цена позиции с учётом скидки' `
    -Formula 'P = P₀ × (1 − d / 100),        (1)' `
    -Legend @('где P — цена с учётом скидки, руб.;', 'P₀ — базовая цена, руб.;', 'd — размер скидки, % (25 %).') `
    -Left 50 -Top 110 -Width 300 -Height 110)

[void](Add-Formula -Slide $s -Title 'Сумма заказа' `
    -Formula 'S = Σ (Pᵢ × qᵢ),        (2)' `
    -Legend @('где S — сумма заказа, руб.;', 'Pᵢ — цена позиции, руб.;', 'qᵢ — количество единиц позиции.') `
    -Left 50 -Top 236 -Width 300 -Height 110)

[void](Add-Text -Slide $s -Text 'Листинг 1 — OrderService.CreateOrderAsync' `
    -Left 380 -Top 100 -Width 530 -Height 22 -FontSize 13 -Bold $false -Color $GRAY)

$code = @(
    '// Скидка 25 % — нет заказов по товару за прошлый месяц',
    'var discount = purchasedItemIds.Contains(itemId)',
    '    ? 0m : DiscountPercent;',
    'var price = item.Price * (1 - discount / 100m);',
    '',
    'item.Quantity -= quantity;',
    '',
    'for (var unit = 0; unit < quantity; unit++)',
    '{',
    '    db.Orders.Add(new Orders {',
    '        Date_Order = now, Item_Id = itemId,',
    '        UserId = userId, Status_Id = status.Id_status });',
    '}',
    'total += price * quantity;'
)
[void](Add-Code -Slide $s -Lines $code -Left 380 -Top 128 -Width 530 -Height 300 -FontSize 14)

# ===== 10. Каталог =====

$s = New-Slide -Title 'Каталог: главная форма приложения' -Note $notes['catalog']

$imgPath = Join-Path $screens '01_main.png'
$pic = Add-Picture -Slide $s -Path $imgPath -Left 50 -Top 100 -MaxWidth 560 -MaxHeight 370

Add-Callout -Slide $s -Pic $pic -ImagePath $imgPath -Rect (Get-Rect 'main.search') -Number 1 | Out-Null
Add-Callout -Slide $s -Pic $pic -ImagePath $imgPath -Rect (Get-Rect 'main.category') -Number 2 | Out-Null
Add-Callout -Slide $s -Pic $pic -ImagePath $imgPath -Rect (Get-Rect 'main.addToCart') -Number 3 | Out-Null
Add-Callout -Slide $s -Pic $pic -ImagePath $imgPath -Rect (Get-Rect 'main.cartButton') -Number 4 | Out-Null
Add-Callout -Slide $s -Pic $pic -ImagePath $imgPath -Rect (Get-Rect 'main.user') -Number 5 -Anchor 'left' | Out-Null

[void](Add-Bullets -Slide $s -Items @(
    '1 — поиск по названию и артикулу',
    '2 — фильтр по категориям',
    '3 — добавление товара в корзину',
    '4 — переход в корзину',
    '5 — текущий пользователь и роль',
    'Поиск и фильтр выполняются одним запросом к базе'
) -Left 640 -Top 120 -Width 270 -Height 340 -FontSize 15)

# ===== 11. Оформление заказа =====

$s = New-Slide -Title 'Оформление заказа: корзина и история заказов' -Note $notes['order']

$imgPath = Join-Path $screens '07_cart.png'
$pic = Add-Picture -Slide $s -Path $imgPath -Left 50 -Top 100 -MaxWidth 280 -MaxHeight 250
Add-Callout -Slide $s -Pic $pic -ImagePath $imgPath -Rect (Get-Rect 'cart.pay') -Number 1 | Out-Null
Add-Callout -Slide $s -Pic $pic -ImagePath $imgPath -Rect (Get-Rect 'cart.total') -Number 2 | Out-Null

$imgPath2 = Join-Path $screens '06_my_orders.png'
$pic2 = Add-Picture -Slide $s -Path $imgPath2 -Left 360 -Top 100 -MaxWidth 300 -MaxHeight 250
Add-Callout -Slide $s -Pic $pic2 -ImagePath $imgPath2 -Rect (Get-Rect 'orders.total') -Number 3 | Out-Null

[void](Add-Bullets -Slide $s -Items @(
    '1 — «Оплатить»: оформляет заказ и списывает остаток',
    '2 — итог корзины с учётом скидки покупателя',
    '3 — количество заказов пользователя',
    '4 — заказ можно отменить, товар возвращается на склад',
    'История заказов доступна только автору'
) -Left 690 -Top 110 -Width 230 -Height 340 -FontSize 15)

# ===== 12. Роли и права =====

$s = New-Slide -Title 'Ролевая модель: роли и разграничение прав' -Note $notes['roles']

[void](Add-Bullets -Slide $s -Items @(
    'Гость — просмотр каталога магазина',
    'Покупатель — оформление и отмена заказа, отзывы, история покупок, профиль',
    'Менеджер — дополнительно обработка заказов: смена статуса и отмена',
    'Администратор — дополнительно управление каталогом, категориями и пользователями',
    'Роль определяется при входе и хранится в таблице Users',
    '9 прав перечислены в Permission, проверка — Permissions.Allows',
    'Проверка выполняется и в интерфейсе, и в сервисе'
) -Top 112)

# ===== 13. Тестирование =====

$s = New-Slide -Title 'Проверка работоспособности продукта' -Note $notes['tests']

[void](Add-Picture -Slide $s -Path (Join-Path $diag 'testrun_output.png') -Left 50 -Top 100 -MaxWidth 560 -MaxHeight 380)

[void](Add-Bullets -Slide $s -Items @(
    '52 модульные проверки сервисов',
    '8 сквозных проверок через UI Automation',
    '0 ошибок, все проверки успешны',
    'Проверены вход и разграничение прав',
    'Проверены скидка 25 % и итоговая сумма',
    'Проверены списание и возврат остатка'
) -Left 640 -Top 120 -Width 270 -Height 340 -FontSize 15)

# ===== 14. Итоги =====

$s = New-Slide -Title 'Итоги работы и направления развития' -Note $notes['result']

[void](Add-Bullets -Slide $s -Items @(
    'Проанализирована предметная область: 11 требований, 4 роли',
    'Спроектирована база: 7 таблиц в 3НФ, 7 внешних ключей, импорт данных',
    'Реализованы архитектура (4 уровня) и алгоритм расчёта цены со скидкой',
    'Разработан интерфейс: 9 форм, корзина, оформление и отмена заказа',
    'Реализована ролевая модель: 4 роли, 9 операций',
    'Проверено: 52 и 8 проверок, 0 ошибок',
    'Развитие: печатные формы, резервное копирование, веб-версия'
) -Top 112)

# ===== 15. Заключительный слайд =====

$s = New-Slide -Note $notes['end']

[void](Add-Text -Slide $s -Text "Санкт-Петербургское государственное автономное`rпрофессиональное образовательное учреждение`r«Колледж туризма и прикладных технологий»" `
    -Left 90 -Top 46 -Width 780 -Height 70 -FontSize 15 -Bold $false -Color $BLACK -Align 2)
[void](Add-Text -Slide $s -Text 'Курсовой проект' -Left 90 -Top 136 -Width 780 -Height 28 -FontSize 16 -Bold $true -Color $BLACK -Align 2)
[void](Add-Text -Slide $s -Text 'на тему:' -Left 90 -Top 170 -Width 780 -Height 24 -FontSize 14 -Bold $false -Color $BLACK -Align 2)
[void](Add-Text -Slide $s -Text "«Разработка информационной системы`rдля магазина автозапчастей»" `
    -Left 90 -Top 196 -Width 780 -Height 64 -FontSize 23 -Bold $true -Color $BLACK -Align 2)
[void](Add-Text -Slide $s -Text 'МДК.01.01 Разработка программных модулей' `
    -Left 90 -Top 268 -Width 780 -Height 24 -FontSize 14 -Bold $false -Color $GRAY -Align 2)
[void](Add-Text -Slide $s -Text "Выполнил(а) студент(ка) гр. 324:`rА. А. Копейкин`rспециальность: 09.02.07`rИнформационные системы и программирование" `
    -Left 110 -Top 322 -Width 380 -Height 90 -FontSize 14 -Bold $false -Color $BLACK -Align 1)
[void](Add-Text -Slide $s -Text "Руководитель:`rТ. Т. Дуларидзе, преподаватель" `
    -Left 520 -Top 322 -Width 330 -Height 60 -FontSize 14 -Bold $false -Color $BLACK -Align 1)
[void](Add-Text -Slide $s -Text 'Санкт-Петербург, 2026' -Left 90 -Top 470 -Width 780 -Height 26 -FontSize 14 -Bold $false -Color $GRAY -Align 2)

# --- Номера страниц ---

$i = 0
foreach ($sl in $slides) {
    $i++
    if ($i -ne 1 -and $i -ne $slides.Count) { Add-PageNumber -Slide $sl -Number $i }
}

if (Test-Path -LiteralPath $outPath) {
    $bak = Join-Path $root ('docs\backup_plain_deck_' + (Get-Date -Format 'yyyy-MM-dd-HHmmss') + '.pptx')
    Copy-Item -LiteralPath $outPath -Destination $bak -Force
}

$pres.SaveAs($outPath, 24)
$pres.Close()
$ppt.Quit()

Write-Output ('OK презентация сохранена: ' + $outPath)
Write-Output ('Слайдов: ' + $slides.Count)
