# -*- coding: utf-8 -*-
# Сборка презентации к защите курсового проекта (12—15 слайдов).
#
# Структура следует регламенту защиты:
#   1        титульный
#   2—3      тема, актуальность, цель и задачи
#   4—6      первая задача: модель данных, ER-диаграмма, результаты импорта
#   7—9      вторая задача: архитектура, алгоритм расчёта цены
#   10—13    третья и четвёртая задачи: интерфейс и ролевая модель
#   14       выводы и направления развития
#   15       заключительный слайд дублирует титульный
#
# Запуск:
#   powershell -File tools/deck/Build-Deck-Defense.ps1

param(
    [string]$ProjectRoot = '.',
    [string]$OutFile = 'Презентация_к_защите_Копейкин.pptx'
)

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'Deck-Lib-Defense.ps1')

$root = (Resolve-Path -LiteralPath $ProjectRoot).Path
$screens = Join-Path $root 'docs\screens'
$diag = Join-Path $root 'docs\diagrams'
$outPath = Join-Path $root $OutFile

$Deck = New-Deck
$Pres = $Deck.Pres

$slides = New-Object System.Collections.ArrayList

function New-Slide {
    param([string]$Id, [string]$Title, [string]$Subtitle = '')
    $s = Add-BlankSlide -Deck $Deck
    [void]$slides.Add([pscustomobject]@{ Slide = $s; Id = $Id; Title = $Title; Notes = '' })
    if ($Title -ne '') { [void](Add-SlideHeader -Slide $s -Title $Title -Subtitle $Subtitle) }
    return $s
}

# Прямоугольники выносок из файла callouts.json.
$calloutFile = Join-Path $screens 'callouts.json'
$callouts = @{}
if (Test-Path -LiteralPath $calloutFile) {
    $callouts = Get-Content -LiteralPath $calloutFile -Raw -Encoding UTF8 | ConvertFrom-Json
}

function Get-Rect {
    param([string]$Key)
    if (-not $callouts.PSObject.Properties.Name.Contains($Key)) { return $null }
    return $callouts.$Key
}

# ============================ 1. Титульный ============================
$s = Add-BlankSlide -Deck $Deck
[void]$slides.Add([pscustomobject]@{ Slide = $s; Id = 'cover'; Title = 'Титульный слайд'; Notes = '' })

$bg = $s.Shapes.AddShape(1, 0, 0, 960, 540)
$bg.Fill.ForeColor.RGB = (New-Rgb 238 242 255)
$bg.Line.Visible = 0
$bg.Shadow.Visible = 0

$head = $s.Shapes.AddShape(1, 0, 0, 960, 96)
$head.Fill.ForeColor.RGB = (New-Rgb 79 70 229)
$head.Line.Visible = 0
$head.Shadow.Visible = 0

$band = $s.Shapes.AddShape(1, 0, 96, 960, 4)
$band.Fill.ForeColor.RGB = (New-Rgb 219 39 119)
$band.Line.Visible = 0
$band.Shadow.Visible = 0

$tb = $s.Shapes.AddTextbox(1, 70, 22, 820, 60)
$tr = $tb.TextFrame.TextRange
$tr.Text = "Санкт-Петербургское государственное автономное`rпрофессиональное образовательное учреждение"
$tr.Font.Name = 'Segoe UI'
$tr.Font.Size = 14
$tr.Font.Bold = -1
$tr.Font.Color.RGB = (New-Rgb 255 255 255)
$tb.TextFrame.WordWrap = -1

$tb = $s.Shapes.AddTextbox(1, 70, 58, 820, 28)
$tr = $tb.TextFrame.TextRange
$tr.Text = '«Колледж туризма и прикладных технологий»'
$tr.Font.Name = 'Segoe UI Semibold'
$tr.Font.Size = 16
$tr.Font.Bold = -1
$tr.Font.Color.RGB = (New-Rgb 255 255 255)
$tb.TextFrame.WordWrap = -1

$tb = $s.Shapes.AddTextbox(1, 70, 126, 820, 26)
$tr = $tb.TextFrame.TextRange
$tr.Text = 'КУРСОВОЙ ПРОЕКТ'
$tr.Font.Name = 'Segoe UI'
$tr.Font.Size = 13
$tr.Font.Bold = -1
$tr.Font.Color.RGB = (New-Rgb 124 58 237)
$tb.TextFrame.WordWrap = -1

$tb = $s.Shapes.AddTextbox(1, 70, 150, 820, 62)
$tr = $tb.TextFrame.TextRange
$tr.Text = 'на тему: «Разработка информационной системы`rдля магазина автозапчастей»'
$tr.Font.Name = 'Segoe UI Semibold'
$tr.Font.Size = 23
$tr.Font.Bold = -1
$tr.Font.Color.RGB = (New-Rgb 17 24 39)
$tb.TextFrame.WordWrap = -1

$tb = $s.Shapes.AddTextbox(1, 70, 222, 820, 24)
$tr = $tb.TextFrame.TextRange
$tr.Text = 'МДК.01.01 Разработка программных модулей'
$tr.Font.Name = 'Segoe UI'
$tr.Font.Size = 14
$tr.Font.Color.RGB = (New-Rgb 100 116 139)
$tb.TextFrame.WordWrap = -1

$ln = $s.Shapes.AddShape(1, 70, 258, 820, 1.2)
$ln.Fill.ForeColor.RGB = (New-Rgb 226 232 240)
$ln.Line.Visible = 0
$ln.Shadow.Visible = 0

$tb = $s.Shapes.AddTextbox(1, 70, 276, 420, 100)
$tr = $tb.TextFrame.TextRange
$tr.Text = "Выполнил(а) студент(ка) гр. 324`rА. А. Копейкин`rспециальность: 09.02.07`rИнформационные системы и программирование"
$tr.Font.Name = 'Segoe UI'
$tr.Font.Size = 13
$tr.Font.Color.RGB = (New-Rgb 17 24 39)
$tb.TextFrame.WordWrap = -1

$tb = $s.Shapes.AddTextbox(1, 520, 276, 370, 60)
$tr = $tb.TextFrame.TextRange
$tr.Text = "Руководитель:`rТ. Т. Дуларидзе, преподаватель"
$tr.Font.Name = 'Segoe UI'
$tr.Font.Size = 13
$tr.Font.Color.RGB = (New-Rgb 17 24 39)
$tb.TextFrame.WordWrap = -1

$tb = $s.Shapes.AddTextbox(1, 70, 486, 820, 26)
$tr = $tb.TextFrame.TextRange
$tr.Text = 'Санкт-Петербург, 2026'
$tr.Font.Name = 'Segoe UI'
$tr.Font.Size = 13
$tr.Font.Color.RGB = (New-Rgb 100 116 139)

$slides[$slides.Count - 1].Notes = 'Здравствуйте! Тема курсового проекта — разработка информационной системы для магазина автозапчастей. Работа выполнена на платформе .NET с использованием Avalonia UI и СУБД SQL Server.'

# ============================ 2. Актуальность и цель ============================
$s = New-Slide -Id 'intro' -Title 'Тема, актуальность и цель работы' -Subtitle 'Постановка задачи курсового проекта'

$cards = @(
    @{ t = 'АКТУАЛЬНОСТЬ'; d = 'Учёт товаров, остатков и заказов ведётся в разных таблицах и бумажных документах — растёт риск ошибок и время обработки заказа'; c = (New-Rgb 219 39 119) },
    @{ t = 'ЦЕЛЬ РАБОТЫ'; d = 'Централизованный учёт ассортимента и складских остатков, автоматизация оформления и сопровождения заказов'; c = (New-Rgb 79 70 229) },
    @{ t = 'ОБЪЕКТ И ПРЕДМЕТ'; d = 'Объект — процесс учёта товаров и заказов в магазине. Предмет — программные средства автоматизации этого процесса'; c = (New-Rgb 2 132 199) }
)
$y = 118
foreach ($c in $cards) {
    [void](Add-Card -Slide $s -Left 50 -Top $y -Width 860 -Height 108 -Fill (New-Rgb 255 255 255))
    $bar = $s.Shapes.AddShape(1, 50, $y, 5, 108)
    $bar.Fill.ForeColor.RGB = $c.c
    $bar.Line.Visible = 0
    $bar.Shadow.Visible = 0
    [void](Add-CardText -Slide $s -Title $c.t -Body $c.d -Left 74 -Top ($y + 12) -Width 818 -Height 84 `
        -TitleSize 13 -BodySize 14 -TitleColor $c.c)
    $y += 122
}
$slides[$slides.Count - 1].Notes = 'Актуальность: данные о товарах и заказах разрозненны, что приводит к ошибкам. Цель — централизованный учёт и автоматизация оформления заказов.'

# ============================ 3. Задачи работы ============================
$s = New-Slide -Id 'tasks' -Title 'Задачи работы и методы проектирования' -Subtitle 'Четыре задачи, решаемые последовательно'

$tasks = @(
    @{ n = '1'; t = 'Проектирование модели данных'; d = 'Нормализация отношений, логическая и физическая модели, импорт данных в SQL Server' },
    @{ n = '2'; t = 'Проектирование архитектуры и алгоритмов'; d = 'Четырёхуровневая архитектура, алгоритм расчёта цены со скидкой и отмены заказа' },
    @{ n = '3'; t = 'Реализация пользовательского интерфейса'; d = 'Девять экранных форм в стиле Fluent, поиск, фильтрация, корзина и оформление заказа' },
    @{ n = '4'; t = 'Разграничение прав доступа и тестирование'; d = 'Четыре роли, девять операций, сквозная проверка через UI Automation' }
)
$y = 112
foreach ($t in $tasks) {
    [void](Add-Card -Slide $s -Left 50 -Top $y -Width 860 -Height 82 -Fill (New-Rgb 255 255 255))
    $num = $s.Shapes.AddShape(9, 66, ($y + 20), 42, 42)
    $num.Fill.ForeColor.RGB = (New-Rgb 79 70 229)
    $num.Line.Visible = 0
    $num.Shadow.Visible = 0
    $num.TextFrame.TextRange.Text = $t.n
    $num.TextFrame.TextRange.Font.Name = 'Segoe UI Semibold'
    $num.TextFrame.TextRange.Font.Size = 18
    $num.TextFrame.TextRange.Font.Bold = -1
    $num.TextFrame.TextRange.Font.Color.RGB = (New-Rgb 255 255 255)
    $num.TextFrame.TextRange.ParagraphFormat.Alignment = 2
    $num.TextFrame.VerticalAnchor = 3
    [void](Add-CardText -Slide $s -Title $t.t -Body $t.d -Left 124 -Top ($y + 12) -Width 768 -Height 62 `
        -TitleSize 15 -BodySize 13 -TitleColor (New-Rgb 17 24 39) -BodyColor (New-Rgb 100 116 139))
    $y += 90
}

$mb = $s.Shapes.AddTextbox(1, 50, 478, 860, 26)
$tr = $mb.TextFrame.TextRange
$tr.Text = 'Методы: анализ документации · функциональное моделирование · нормализация отношений · объектно-ориентированное проектирование · модульное тестирование'
$tr.Font.Name = 'Segoe UI'
$tr.Font.Size = 12
$tr.Font.Italic = -1
$tr.Font.Color.RGB = (New-Rgb 100 116 139)
$mb.TextFrame.WordWrap = -1
$slides[$slides.Count - 1].Notes = 'Работа решалась последовательно: сначала модель данных, затем архитектура и алгоритмы, далее интерфейс и разграничение прав с тестированием.'

# ============================ 4. ER-диаграмма ============================
$s = New-Slide -Id 'er' -Title 'Первая задача: логическая модель данных' -Subtitle 'ER-диаграмма базы данных — семь таблиц в третьей нормальной форме'

$pic = Add-PictureFit -Slide $s -Path (Join-Path $diag 'er.png') -Left 44 -Top 112 -MaxWidth 640 -MaxHeight 372
[void](Add-Caption -Slide $s -Text 'Рисунок 1 — ER-диаграмма базы данных' -Left 44 -Top 490 -Width 640)

$legend = @(
    'Справочники: Roles, Category, OrderStatus',
    'Сущности: Users, Item, Orders, ItemReview',
    '7 внешних ключей, устранение аномалий',
    '43 записи в справочниках при старте'
)
$y = 118
foreach ($l in $legend) {
    [void](Add-Card -Slide $s -Left 700 -Top $y -Width 216 -Height 74 -Fill (New-Rgb 255 255 255))
    [void](Add-Label -Slide $s -Text $l -Left 714 -Top ($y + 12) -Width 190 -Height 52 -Size 12 -Bold $false -Color (New-Rgb 17 24 39))
    $y += 84
}
$slides[$slides.Count - 1].Notes = 'Модель данных приведена к третьей нормальной форме: справочники ролей, категорий и статусов отделены от операционных таблиц. Всего семь таблиц, семь внешних ключей.'

# ============================ 5. Физическая модель и импорт ============================
$s = New-Slide -Id 'physical' -Title 'Первая задача: физическая модель и наполнение базы' -Subtitle 'Скрипт создания базы на SQL Server и результаты импорта данных'

$rows = @(
    @('Таблица', 'Назначение', 'Строк'),
    @('Roles', 'Справочник ролей пользователей', '4'),
    @('Users', 'Учётные записи покупателей и сотрудников', '5'),
    @('Category', 'Категории (отделения) автозапчастей', '6'),
    @('Item', 'Каталог товаров: цена, остаток, описание', '12'),
    @('Orders', 'Строки заказов со статусом и датой', '12'),
    @('OrderStatus', 'Справочник статусов заказа', '6'),
    @('ItemReview', 'Отзывы покупателей о товарах', '15')
)
[void](Add-Table -Slide $s -Rows $rows -Left 50 -Top 116 -Width 560 -Height 300 -FontSize 12 -ColWidths @(120, 330, 110))

[void](Add-Card -Slide $s -Left 640 -Top 116 -Width 270 -Height 300 -Fill (New-Rgb 255 255 255))
[void](Add-Label -Slide $s -Text 'РЕЗУЛЬТАТ ИМПОРТА' -Left 658 -Top 132 -Width 236 -Height 22 -Size 12 -Bold $true -Color (New-Rgb 79 70 229))
[void](Add-Label -Slide $s -Text "База создаётся скриптом`rSQLCreateTable.sql,`rданные загружаются`rinsert_sqlserver.sql`r`rКонтроль целостности:`rвнешние ключи и`rограничения NOT NULL" -Left 658 -Top 160 -Width 238 -Height 200 -Size 13 -Bold $false -Color (New-Rgb 17 24 39))

[void](Add-Kpi -Slide $s -Value '7' -Caption 'таблиц в базе' -Left 50 -Top 428 -Width 170 -Height 74 -Color (New-Rgb 79 70 229))
[void](Add-Kpi -Slide $s -Value '7' -Caption 'внешних ключей' -Left 236 -Top 428 -Width 170 -Height 74 -Color (New-Rgb 124 58 237))
[void](Add-Kpi -Slide $s -Value '43' -Caption 'записи справочников' -Left 422 -Top 428 -Width 170 -Height 74 -Color (New-Rgb 219 39 119))
[void](Add-Kpi -Slide $s -Value '12' -Caption 'позиций каталога' -Left 608 -Top 428 -Width 170 -Height 74 -Color (New-Rgb 5 150 105))
[void](Add-Kpi -Slide $s -Value '1' -Caption 'база cursevork' -Left 794 -Top 428 -Width 116 -Height 74 -Color (New-Rgb 2 132 199))

$slides[$slides.Count - 1].Notes = 'Физическая модель реализована на SQL Server. Скрипты создания таблиц и загрузки данных лежат в папке sql. Импорт выполнен без ошибок, все ограничения целостности соблюдены.'

# ============================ 6. Архитектура ============================
$s = New-Slide -Id 'architecture' -Title 'Вторая задача: архитектура приложения' -Subtitle 'Четыре уровня, каждый обращается только к следующему'

$pic = Add-PictureFit -Slide $s -Path (Join-Path $diag 'architecture.png') -Left 44 -Top 108 -MaxWidth 560 -MaxHeight 376
[void](Add-Caption -Slide $s -Text 'Рисунок 2 — Архитектура информационной системы' -Left 44 -Top 490 -Width 560)

$levels = @(
    @{ t = 'Интерфейс пользователя'; d = 'Avalonia UI, девять экранных форм на XAML'; c = (New-Rgb 219 39 119) },
    @{ t = 'Логика окон'; d = 'Модели представления (ViewModels), шаблон MVVM'; c = (New-Rgb 124 58 237) },
    @{ t = 'Бизнес-правила'; d = 'Сервисы: заказы, каталог, отзывы, роли'; c = (New-Rgb 79 70 229) },
    @{ t = 'Доступ к данным'; d = 'Контекст EF Core, фабрика DbFactory'; c = (New-Rgb 2 132 199) }
)
$y = 112
foreach ($l in $levels) {
    [void](Add-Card -Slide $s -Left 624 -Top $y -Width 286 -Height 78 -Fill (New-Rgb 255 255 255))
    $bar = $s.Shapes.AddShape(1, 624, $y, 4, 78)
    $bar.Fill.ForeColor.RGB = $l.c
    $bar.Line.Visible = 0
    $bar.Shadow.Visible = 0
    [void](Add-CardText -Slide $s -Title $l.t -Body $l.d -Left 640 -Top ($y + 11) -Width 256 -Height 58 `
        -TitleSize 13 -BodySize 12 -TitleColor $l.c -BodyColor (New-Rgb 100 116 139))
    $y += 88
}
$slides[$slides.Count - 1].Notes = 'Архитектура четырёхуровневая: представление не обращается к базе напрямую, а идёт через модель представления и сервисы. Это упрощает тестирование и сопровождение.'

# ============================ 7. Алгоритм расчёта цены ============================
$s = New-Slide -Id 'algorithm' -Title 'Вторая задача: алгоритм оформления заказа' -Subtitle 'Схема по ГОСТ 19.701-90 и правило расчёта цены со скидкой'

$pic = Add-PictureFit -Slide $s -Path (Join-Path $diag 'alg_create.png') -Left 44 -Top 108 -MaxWidth 520 -MaxHeight 350
[void](Add-Caption -Slide $s -Text 'Рисунок 3 — Схема алгоритма оформления заказа' -Left 44 -Top 464 -Width 520)

# Формулы
$fbox = Add-Card -Slide $s -Left 584 -Top 106 -Width 326 -Height 206 -Fill (New-Rgb 255 255 255)
[void](Add-Label -Slide $s -Text 'Цена позиции с учётом скидки' -Left 600 -Top 116 -Width 296 -Height 20 -Size 12 -Bold $false -Color (New-Rgb 100 116 139))
[void](Add-Label -Slide $s -Text 'P = P₀ × (1 − d / 100),                    (1)' -Left 600 -Top 136 -Width 296 -Height 24 -Size 15 -Bold $true -Color (New-Rgb 17 24 39))
[void](Add-Label -Slide $s -Text "где P — цена с учётом скидки, руб.;`rP₀ — базовая цена, руб.;`rd — размер скидки, % (25 %)." -Left 600 -Top 164 -Width 296 -Height 46 -Size 12 -Bold $false -Color (New-Rgb 17 24 39))
[void](Add-Label -Slide $s -Text 'Сумма заказа' -Left 600 -Top 214 -Width 296 -Height 20 -Size 12 -Bold $false -Color (New-Rgb 100 116 139))
[void](Add-Label -Slide $s -Text 'S = Σ (Pᵢ × qᵢ),                              (2)' -Left 600 -Top 234 -Width 296 -Height 24 -Size 15 -Bold $true -Color (New-Rgb 17 24 39))
[void](Add-Label -Slide $s -Text "где S — сумма заказа, руб.;`rPᵢ — цена позиции с учётом скидки, руб.;`rqᵢ — количество единиц позиции." -Left 600 -Top 262 -Width 296 -Height 46 -Size 12 -Bold $false -Color (New-Rgb 17 24 39))

[void](Add-Card -Slide $s -Left 584 -Top 316 -Width 326 -Height 152 -Fill (New-Rgb 255 255 255))
[void](Add-Label -Slide $s -Text 'ПРАВИЛА РАБОТЫ АЛГОРИТМА' -Left 600 -Top 328 -Width 296 -Height 20 -Size 12 -Bold $true -Color (New-Rgb 79 70 229))
$rules = @(
    'Скидка 25 % — нет заказа по товару за прошлый месяц',
    'Остаток списывается в той же транзакции, что и заказ',
    'Количество разворачивается в отдельные строки заказа',
    'Заказ свыше пяти позиций помечается как крупный'
)
$ry = 352
foreach ($r in $rules) {
    [void](Add-Label -Slide $s -Text "• $r" -Left 600 -Top $ry -Width 298 -Height 28 -Size 11.5 -Bold $false -Color (New-Rgb 17 24 39))
    $ry += 29
}
$slides[$slides.Count - 1].Notes = 'Ключевое правило: скидка двадцать пять процентов даётся, если в прошлом календарном месяце по этому товару заказа не было. Формула (1) даёт цену позиции, формула (2) — сумму заказа.'

# ============================ 8. Фрагмент кода ============================
$s = New-Slide -Id 'code' -Title 'Вторая задача: реализация расчёта цены' -Subtitle 'Фрагмент метода OrderService.CreateOrderAsync'

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
    '',
    'total += price * quantity;'
)
[void](Add-CodeBlock -Slide $s -Lines $code -Left 50 -Top 112 -Width 570 -Height 300 -FontSize 14)
[void](Add-Caption -Slide $s -Text 'Листинг 1 — расчёт цены позиции и списание остатка' -Left 50 -Top 418 -Width 570)

[void](Add-Card -Slide $s -Left 650 -Top 112 -Width 260 -Height 300 -Fill (New-Rgb 255 255 255))
[void](Add-Label -Slide $s -Text 'ПОЧЕМУ ТАК' -Left 666 -Top 128 -Width 228 -Height 20 -Size 12 -Bold $true -Color (New-Rgb 79 70 229))
$why = @(
    'Остаток на экране не источник истины — он проверяется в базе',
    'Транзакция не даёт записать заказ при нулевом остатке',
    'Скидка считается по истории заказов, а не по полю в таблице',
    'Каждая единица товара — отдельная строка Orders'
)
$wy = 156
foreach ($w in $why) {
    [void](Add-Label -Slide $s -Text "• $w" -Left 666 -Top $wy -Width 232 -Height 50 -Size 12 -Bold $false -Color (New-Rgb 17 24 39))
    $wy += 64
}
$slides[$slides.Count - 1].Notes = 'Здесь показан ключевой фрагмент. Обратите внимание: остаток списывается внутри транзакции, а скидка определяется по истории заказов пользователя за прошлый месяц.'

# ============================ 9. Главное окно ============================
$s = New-Slide -Id 'ui-main' -Title 'Третья задача: главная форма приложения' -Subtitle 'Каталог автозапчастей с поиском, фильтрацией и корзиной'

$imgPath = Join-Path $screens '01_main.png'
$pic = Add-PictureFit -Slide $s -Path $imgPath -Left 44 -Top 108 -MaxWidth 560 -MaxHeight 340

Add-Callout -Slide $s -Pic $pic -ImagePath $imgPath -Rect (Get-Rect 'main.search') -Number 1 -Anchor 'tr' | Out-Null
Add-Callout -Slide $s -Pic $pic -ImagePath $imgPath -Rect (Get-Rect 'main.category') -Number 2 -Anchor 'tr' | Out-Null
Add-Callout -Slide $s -Pic $pic -ImagePath $imgPath -Rect (Get-Rect 'main.addToCart') -Number 3 -Anchor 'tr' | Out-Null
Add-Callout -Slide $s -Pic $pic -ImagePath $imgPath -Rect (Get-Rect 'main.cartButton') -Number 4 -Anchor 'tr' | Out-Null
Add-Callout -Slide $s -Pic $pic -ImagePath $imgPath -Rect (Get-Rect 'main.user') -Number 5 -Anchor 'left' | Out-Null

[void](Add-Caption -Slide $s -Text 'Рисунок 4 — Главная форма приложения' -Left 44 -Top 456 -Width 560)

[void](Add-Card -Slide $s -Left 624 -Top 108 -Width 286 -Height 348 -Fill (New-Rgb 255 255 255))
$legend = @(
    '1 — поиск по названию и артикулу',
    '2 — фильтр по категории',
    '3 — добавление товара в корзину',
    '4 — переход в корзину',
    '5 — текущий пользователь и роль'
)
$y = 126
$i = 0
foreach ($l in $legend) {
    $dot = $s.Shapes.AddShape(9, 642, ($y + 2), 18, 18)
    $dot.Fill.ForeColor.RGB = (New-Rgb 219 39 119)
    $dot.Line.Visible = 0
    $dot.Shadow.Visible = 0
    $dot.TextFrame.TextRange.Text = "$($i + 1)"
    $dot.TextFrame.TextRange.Font.Name = 'Segoe UI'
    $dot.TextFrame.TextRange.Font.Size = 10
    $dot.TextFrame.TextRange.Font.Bold = -1
    $dot.TextFrame.TextRange.Font.Color.RGB = (New-Rgb 255 255 255)
    $dot.TextFrame.TextRange.ParagraphFormat.Alignment = 2
    $dot.TextFrame.VerticalAnchor = 3
    [void](Add-Label -Slide $s -Text $l.Substring(4) -Left 668 -Top $y -Width 230 -Height 40 -Size 12 -Bold $false -Color (New-Rgb 17 24 39))
    $y += 52
    $i++
}
[void](Add-Label -Slide $s -Text 'Поиск и фильтр применяются одновременно — выборка выполняется одним запросом к базе.' -Left 642 -Top 392 -Width 254 -Height 52 -Size 11 -Bold $false -Color (New-Rgb 100 116 139))
$slides[$slides.Count - 1].Notes = 'Главная форма: каталог с фотографиями, поиск по названию и артикулу, фильтр по категориям. Поиск и фильтр работают одновременно и выполняются на стороне базы.'

# ============================ 10. Корзина и заказы ============================
$s = New-Slide -Id 'ui-cart' -Title 'Третья задача: оформление заказа' -Subtitle 'Корзина покупателя и история собственных заказов'

$imgPath = Join-Path $screens '07_cart.png'
$pic = Add-PictureFit -Slide $s -Path $imgPath -Left 44 -Top 112 -MaxWidth 400 -MaxHeight 280
Add-Callout -Slide $s -Pic $pic -ImagePath $imgPath -Rect (Get-Rect 'cart.pay') -Number 1 -Anchor 'tr' | Out-Null
Add-Callout -Slide $s -Pic $pic -ImagePath $imgPath -Rect (Get-Rect 'cart.total') -Number 2 -Anchor 'tr' | Out-Null
[void](Add-Caption -Slide $s -Text 'Рисунок 5 — Корзина покупателя' -Left 44 -Top 400 -Width 400)

$imgPath2 = Join-Path $screens '06_my_orders.png'
$pic2 = Add-PictureFit -Slide $s -Path $imgPath2 -Left 480 -Top 112 -MaxWidth 430 -MaxHeight 280
Add-Callout -Slide $s -Pic $pic2 -ImagePath $imgPath2 -Rect (Get-Rect 'orders.total') -Number 3 -Anchor 'tr' | Out-Null
[void](Add-Caption -Slide $s -Text 'Рисунок 6 — История заказов покупателя' -Left 480 -Top 400 -Width 430)

[void](Add-Card -Slide $s -Left 44 -Top 428 -Width 866 -Height 76 -Fill (New-Rgb 255 255 255))
$notes = @(
    @{ n = '1'; d = 'Кнопка «Оплатить» оформляет заказ и списывает остаток' },
    @{ n = '2'; d = 'Итоговая сумма корзины с учётом скидки постоянного покупателя' },
    @{ n = '3'; d = 'История заказов доступна только автору, заказ можно отменить' }
)
$nx = 64
foreach ($n in $notes) {
    $dot = $s.Shapes.AddShape(9, $nx, 450, 18, 18)
    $dot.Fill.ForeColor.RGB = (New-Rgb 219 39 119)
    $dot.Line.Visible = 0
    $dot.Shadow.Visible = 0
    $dot.TextFrame.TextRange.Text = $n.n
    $dot.TextFrame.TextRange.Font.Name = 'Segoe UI'
    $dot.TextFrame.TextRange.Font.Size = 10
    $dot.TextFrame.TextRange.Font.Bold = -1
    $dot.TextFrame.TextRange.Font.Color.RGB = (New-Rgb 255 255 255)
    $dot.TextFrame.TextRange.ParagraphFormat.Alignment = 2
    $dot.TextFrame.VerticalAnchor = 3
    [void](Add-Label -Slide $s -Text $n.d -Left ($nx + 24) -Top 447 -Width 258 -Height 42 -Size 11 -Bold $false -Color (New-Rgb 17 24 39))
    $nx += 288
}
$slides[$slides.Count - 1].Notes = 'Корзина: количество позиций меняется кнопками плюс и минус, удаление возвращает товар в каталог. Оплата оформляет заказ, номер и сумма показываются сразу. Отмена заказа возвращает товар на склад.'

# ============================ 11. Обработка заказов и управление ============================
$s = New-Slide -Id 'ui-manager' -Title 'Четвёртая задача: ролевая модель в интерфейсе' -Subtitle 'Сотрудник сопровождает заказ, администратор управляет магазином'

$imgPath = Join-Path $screens '04_manager_orders.png'
$pic = Add-PictureFit -Slide $s -Path $imgPath -Left 44 -Top 112 -MaxWidth 420 -MaxHeight 270
Add-Callout -Slide $s -Pic $pic -ImagePath $imgPath -Rect (Get-Rect 'manager.statusCombo') -Number 1 -Anchor 'tr' | Out-Null
Add-Callout -Slide $s -Pic $pic -ImagePath $imgPath -Rect (Get-Rect 'manager.cancel') -Number 2 -Anchor 'tr' | Out-Null
[void](Add-Caption -Slide $s -Text 'Рисунок 7 — Обработка заказов сотрудником' -Left 44 -Top 390 -Width 420)

$imgPath2 = Join-Path $screens '05_admin.png'
$pic2 = Add-PictureFit -Slide $s -Path $imgPath2 -Left 490 -Top 112 -MaxWidth 420 -MaxHeight 270
[void](Add-Caption -Slide $s -Text 'Рисунок 8 — Управление магазином администратором' -Left 490 -Top 390 -Width 420)

[void](Add-Card -Slide $s -Left 44 -Top 418 -Width 866 -Height 86 -Fill (New-Rgb 255 255 255))
[void](Add-Label -Slide $s -Text '1 — выбор нового статуса заказа   ·   2 — отмена заказа с возвратом товара на склад' -Left 62 -Top 430 -Width 830 -Height 22 -Size 12 -Bold $false -Color (New-Rgb 17 24 39))
[void](Add-Label -Slide $s -Text 'Кнопки личных разделов скрываются по роли: сотрудник не видит управления каталогом, покупатель — чужих заказов. Скрытие кнопки не заменяет проверку права в сервисе.' -Left 62 -Top 456 -Width 832 -Height 44 -Size 12 -Bold $false -Color (New-Rgb 100 116 139))
$slides[$slides.Count - 1].Notes = 'Сотрудник видит все заказы магазина и меняет их статус, администратор — ещё и управляет каталогом и пользователями. Обратите внимание: проверка права есть и в интерфейсе, и в сервисе.'

# ============================ 12. Матрица прав ============================
$s = New-Slide -Id 'roles' -Title 'Четвёртая задача: разграничение прав доступа' -Subtitle 'Четыре роли, девять операций, проверка в интерфейсе и в сервисе'

$rows = @(
    @('Операция', 'Гость', 'Покупатель', 'Менеджер', 'Админ'),
    @('Просмотр каталога', '+', '+', '+', '+'),
    @('Оформление заказа', '−', '+', '+', '+'),
    @('История своих заказов', '−', '+', '+', '+'),
    @('Изменение профиля', '−', '+', '+', '+'),
    @('Отзывы о товарах', '−', '+', '+', '+'),
    @('Обработка заказов', '−', '−', '+', '+'),
    @('Управление каталогом', '−', '−', '−', '+'),
    @('Управление пользователями', '−', '−', '−', '+')
)
[void](Add-Table -Slide $s -Rows $rows -Left 50 -Top 114 -Width 580 -Height 330 -FontSize 12 -ColWidths @(240, 85, 90, 90, 75))

[void](Add-Card -Slide $s -Left 650 -Top 114 -Width 260 -Height 330 -Fill (New-Rgb 255 255 255))
[void](Add-Label -Slide $s -Text 'КАК РЕАЛИЗОВАНО' -Left 666 -Top 130 -Width 228 -Height 20 -Size 12 -Bold $true -Color (New-Rgb 79 70 229))
$how = @(
    'Перечисление Permission — девять прав',
    'Метод Permissions.Allows(role, permission)',
    'Роль определяется при входе и хранится в Users',
    'Проверка и в интерфейсе, и в сервисе'
)
$hy = 158
foreach ($h in $how) {
    [void](Add-Label -Slide $s -Text "• $h" -Left 666 -Top $hy -Width 232 -Height 46 -Size 12 -Bold $false -Color (New-Rgb 17 24 39))
    $hy += 58
}
[void](Add-Label -Slide $s -Text '«+» — доступно, «−» — недоступно' -Left 666 -Top 396 -Width 232 -Height 22 -Size 11 -Bold $false -Color (New-Rgb 100 116 139))
$slides[$slides.Count - 1].Notes = 'Матрица прав. Гость только просматривает каталог. Покупатель оформляет заказы и пишет отзывы. Менеджер обрабатывает заказы. Администратор управляет всем.'

# ============================ 13. Тестирование ============================
$s = New-Slide -Id 'testing' -Title 'Проверка работоспособности продукта' -Subtitle 'Сквозная проверка через UI Automation и модульные тесты сервисов'

$pic = Add-PictureFit -Slide $s -Path (Join-Path $diag 'testrun_output.png') -Left 44 -Top 112 -MaxWidth 570 -MaxHeight 320
[void](Add-Caption -Slide $s -Text 'Рисунок 9 — Протокол сквозной проверки покупки' -Left 44 -Top 440 -Width 570)

[void](Add-Kpi -Slide $s -Value '52' -Caption 'проверок сервисов' -Left 640 -Top 112 -Width 130 -Height 76 -Color (New-Rgb 79 70 229))
[void](Add-Kpi -Slide $s -Value '8' -Caption 'проверок через UI' -Left 786 -Top 112 -Width 124 -Height 76 -Color (New-Rgb 124 58 237))
[void](Add-Kpi -Slide $s -Value '0' -Caption 'ошибок' -Left 640 -Top 200 -Width 130 -Height 76 -Color (New-Rgb 5 150 105))

[void](Add-Card -Slide $s -Left 786 -Top 200 -Width 124 -Height 76 -Fill (New-Rgb 255 255 255))
[void](Add-Label -Slide $s -Text '4' -Left 798 -Top 208 -Width 100 -Height 30 -Size 26 -Bold $true -Color (New-Rgb 2 132 199))
$kb = Add-Label -Slide $s -Text 'роли проверены' -Left 796 -Top 240 -Width 104 -Height 24 -Size 12 -Bold $false -Color (New-Rgb 100 116 139)
$kb.TextFrame.TextRange.ParagraphFormat.Alignment = 2

[void](Add-Card -Slide $s -Left 640 -Top 290 -Width 270 -Height 174 -Fill (New-Rgb 255 255 255))
[void](Add-Label -Slide $s -Text 'ЧТО ПРОВЕРЕНО' -Left 656 -Top 304 -Width 238 -Height 20 -Size 12 -Bold $true -Color (New-Rgb 79 70 229))
$checked = @(
    'Вход и разграничение прав',
    'Скидка 25 % и расчёт суммы',
    'Списание и возврат остатка'
)
$cy = 332
foreach ($c in $checked) {
    [void](Add-Label -Slide $s -Text "• $c" -Left 656 -Top $cy -Width 240 -Height 30 -Size 12 -Bold $false -Color (New-Rgb 17 24 39))
    $cy += 36
}
$slides[$slides.Count - 1].Notes = 'Проверка выполнена в двух видах: модульные тесты сервисов — пятьдесят две проверки, и сквозная проверка через UI Automation — восемь проверок. Все завершились успешно.'

# ============================ 14. Выводы ============================
$s = New-Slide -Id 'conclusion' -Title 'Выводы и направления развития системы' -Subtitle 'Достигнутые результаты и перспективы доработки'

$done = @(
    @{ t = 'Проанализирована предметная область'; d = 'Выделены бизнес-процессы, 11 функциональных требований, 4 роли участников' },
    @{ t = 'Спроектирована и реализована база данных'; d = '7 таблиц в третьей нормальной форме, 7 внешних ключей, импорт данных в SQL Server' },
    @{ t = 'Реализованы архитектура и алгоритмы'; d = 'Четырёхуровневая архитектура, алгоритм расчёта цены со скидкой и отмены заказа' },
    @{ t = 'Разработан интерфейс и ролевая модель'; d = '9 экранных форм, 4 роли, 9 операций, проверка через UI Automation' }
)
$y = 112
foreach ($d in $done) {
    [void](Add-Card -Slide $s -Left 50 -Top $y -Width 560 -Height 76 -Fill (New-Rgb 255 255 255))
    $bar = $s.Shapes.AddShape(1, 50, $y, 4, 76)
    $bar.Fill.ForeColor.RGB = (New-Rgb 5 150 105)
    $bar.Line.Visible = 0
    $bar.Shadow.Visible = 0
    [void](Add-CardText -Slide $s -Title $d.t -Body $d.d -Left 70 -Top ($y + 11) -Width 528 -Height 56 `
        -TitleSize 13 -BodySize 11 -TitleColor (New-Rgb 17 24 39) -BodyColor (New-Rgb 100 116 139))
    $y += 84
}

[void](Add-Card -Slide $s -Left 630 -Top 112 -Width 280 -Height 328 -Fill (New-Rgb 255 255 255))
[void](Add-Label -Slide $s -Text 'НАПРАВЛЕНИЯ РАЗВИТИЯ' -Left 646 -Top 128 -Width 248 -Height 20 -Size 12 -Bold $true -Color (New-Rgb 219 39 119))
$future = @(
    'Печатные формы накладной и чека',
    'Резервное копирование базы',
    'Веб-версия приложения'
)
$fy = 168
foreach ($f in $future) {
    [void](Add-Label -Slide $s -Text "• $f" -Left 646 -Top $fy -Width 250 -Height 44 -Size 12 -Bold $false -Color (New-Rgb 17 24 39))
    $fy += 58
}
$slides[$slides.Count - 1].Notes = 'Все поставленные задачи решены: модель данных, архитектура и алгоритмы, интерфейс и ролевая модель. Перспективы — печатные формы, резервное копирование и веб-версия.'

# ============================ 15. Заключительный слайд ============================
$s = Add-BlankSlide -Deck $Deck
[void]$slides.Add([pscustomobject]@{ Slide = $s; Id = 'end'; Title = 'Заключительный слайд'; Notes = '' })

$bg = $s.Shapes.AddShape(1, 0, 0, 960, 540)
$bg.Fill.ForeColor.RGB = (New-Rgb 238 242 255)
$bg.Line.Visible = 0
$bg.Shadow.Visible = 0

$head = $s.Shapes.AddShape(1, 0, 0, 960, 96)
$head.Fill.ForeColor.RGB = (New-Rgb 79 70 229)
$head.Line.Visible = 0
$head.Shadow.Visible = 0

$band = $s.Shapes.AddShape(1, 0, 96, 960, 4)
$band.Fill.ForeColor.RGB = (New-Rgb 219 39 119)
$band.Line.Visible = 0
$band.Shadow.Visible = 0

$tb = $s.Shapes.AddTextbox(1, 70, 22, 820, 60)
$tr = $tb.TextFrame.TextRange
$tr.Text = "Санкт-Петербургское государственное автономное`rпрофессиональное образовательное учреждение"
$tr.Font.Name = 'Segoe UI'
$tr.Font.Size = 14
$tr.Font.Bold = -1
$tr.Font.Color.RGB = (New-Rgb 255 255 255)
$tb.TextFrame.WordWrap = -1

$tb = $s.Shapes.AddTextbox(1, 70, 58, 820, 28)
$tr = $tb.TextFrame.TextRange
$tr.Text = '«Колледж туризма и прикладных технологий»'
$tr.Font.Name = 'Segoe UI Semibold'
$tr.Font.Size = 16
$tr.Font.Bold = -1
$tr.Font.Color.RGB = (New-Rgb 255 255 255)
$tb.TextFrame.WordWrap = -1

$tb = $s.Shapes.AddTextbox(1, 70, 126, 820, 26)
$tr = $tb.TextFrame.TextRange
$tr.Text = 'КУРСОВОЙ ПРОЕКТ'
$tr.Font.Name = 'Segoe UI'
$tr.Font.Size = 13
$tr.Font.Bold = -1
$tr.Font.Color.RGB = (New-Rgb 124 58 237)
$tb.TextFrame.WordWrap = -1

$tb = $s.Shapes.AddTextbox(1, 70, 150, 820, 62)
$tr = $tb.TextFrame.TextRange
$tr.Text = 'на тему: «Разработка информационной системы`rдля магазина автозапчастей»'
$tr.Font.Name = 'Segoe UI Semibold'
$tr.Font.Size = 23
$tr.Font.Bold = -1
$tr.Font.Color.RGB = (New-Rgb 17 24 39)
$tb.TextFrame.WordWrap = -1

$tb = $s.Shapes.AddTextbox(1, 70, 222, 820, 24)
$tr = $tb.TextFrame.TextRange
$tr.Text = 'МДК.01.01 Разработка программных модулей'
$tr.Font.Name = 'Segoe UI'
$tr.Font.Size = 14
$tr.Font.Color.RGB = (New-Rgb 100 116 139)
$tb.TextFrame.WordWrap = -1

$ln = $s.Shapes.AddShape(1, 70, 258, 820, 1.2)
$ln.Fill.ForeColor.RGB = (New-Rgb 226 232 240)
$ln.Line.Visible = 0
$ln.Shadow.Visible = 0

$tb = $s.Shapes.AddTextbox(1, 70, 276, 420, 100)
$tr = $tb.TextFrame.TextRange
$tr.Text = "Выполнил(а) студент(ка) гр. 324`rА. А. Копейкин`rспециальность: 09.02.07`rИнформационные системы и программирование"
$tr.Font.Name = 'Segoe UI'
$tr.Font.Size = 13
$tr.Font.Color.RGB = (New-Rgb 17 24 39)
$tb.TextFrame.WordWrap = -1

$tb = $s.Shapes.AddTextbox(1, 520, 276, 370, 60)
$tr = $tb.TextFrame.TextRange
$tr.Text = "Руководитель:`rТ. Т. Дуларидзе, преподаватель"
$tr.Font.Name = 'Segoe UI'
$tr.Font.Size = 13
$tr.Font.Color.RGB = (New-Rgb 17 24 39)
$tb.TextFrame.WordWrap = -1

$tb = $s.Shapes.AddTextbox(1, 70, 486, 820, 26)
$tr = $tb.TextFrame.TextRange
$tr.Text = 'Санкт-Петербург, 2026'
$tr.Font.Name = 'Segoe UI'
$tr.Font.Size = 13
$tr.Font.Color.RGB = (New-Rgb 100 116 139)

$slides[$slides.Count - 1].Notes = 'Доклад окончен. Благодарю за внимание, готов(а) ответить на вопросы.'

# ============================ Финализация ============================
$total = $slides.Count
$i = 0
foreach ($sl in $slides) {
    $i++
    if ($sl.Id -ne 'cover' -and $sl.Id -ne 'end') {
        Add-PageNumber -Slide $sl.Slide -Number $i -Total "$total" | Out-Null
        Add-Footer -Slide $sl.Slide | Out-Null
    }
    if ($sl.Notes -ne '') { Add-SpeakerNote -Slide $sl.Slide -Text $sl.Notes }
}

if (Test-Path -LiteralPath $outPath) {
    $bak = Join-Path $root ('docs\backup_deck_defense_' + (Get-Date -Format 'yyyy-MM-dd-HHmmss') + '.pptx')
    Copy-Item -LiteralPath $outPath -Destination $bak -Force
}

$Pres.SaveAs($outPath, 24)   # ppSaveAsOpenXMLPresentation
$Pres.Close()
$Deck.App.Quit()

Write-Output ('OK презентация сохранена: ' + $outPath)
Write-Output ('Слайдов: ' + $total)
