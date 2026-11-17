# -*- coding: utf-8 -*-
# Сборка презентации по пояснительной записке курсового проекта.
param(
    [string]$ProjectRoot = '.',
    [string]$OutFile = 'Презентация_курсовой_проект_Копейкин.pptx'
)

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'Deck-Lib.ps1')

$root = (Resolve-Path -LiteralPath $ProjectRoot).Path
$diag = Join-Path $root 'docs\diagrams'
$scr = Join-Path $root 'docs\screens'
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

# ============================ 1. Титульный ============================
$s = Add-BlankSlide -Deck $Deck
[void]$slides.Add([pscustomobject]@{ Slide = $s; Id = 'cover'; Title = 'Титульный слайд'; Notes = '' })

$bg = $s.Shapes.AddShape(1, 0, 0, 960, 540)
$bg.Fill.ForeColor.RGB = (New-Rgb 31 56 100)
$bg.Line.Visible = 0
$bg.Shadow.Visible = 0

$band = $s.Shapes.AddShape(1, 0, 396, 960, 4)
$band.Fill.ForeColor.RGB = (New-Rgb 197 90 17)
$band.Line.Visible = 0
$band.Shadow.Visible = 0

$tb = $s.Shapes.AddTextbox(1, 70, 128, 820, 40)
$tr = $tb.TextFrame.TextRange
$tr.Text = 'МИНИСТЕРСТВО ОБРАЗОВАНИЯ И НАУКИ РОССИЙСКОЙ ФЕДЕРАЦИИ'
$tr.Font.Name = 'Segoe UI'
$tr.Font.Size = 12
$tr.Font.Color.RGB = (New-Rgb 170 195 230)
$tb.TextFrame.WordWrap = -1

$tb = $s.Shapes.AddTextbox(1, 70, 176, 820, 110)
$tr = $tb.TextFrame.TextRange
$tr.Text = 'Разработка информационной системы для магазина автозапчастей'
$tr.Font.Name = 'Segoe UI Semibold'
$tr.Font.Size = 34
$tr.Font.Bold = -1
$tr.Font.Color.RGB = (New-Rgb 255 255 255)
$tb.TextFrame.WordWrap = -1

$tb = $s.Shapes.AddTextbox(1, 70, 300, 820, 70)
$tr = $tb.TextFrame.TextRange
$tr.Text = 'Курсовой проект по дисциплине «Информационные технологии и системы»'
$tr.Font.Name = 'Segoe UI'
$tr.Font.Size = 16
$tr.Font.Color.RGB = (New-Rgb 200 218 240)
$tb.TextFrame.WordWrap = -1

$tb = $s.Shapes.AddTextbox(1, 70, 424, 400, 80)
$tr = $tb.TextFrame.TextRange
$tr.Text = "Выполнил: студент группы ИС-41`rКопейкин Андрей"
$tr.Font.Name = 'Segoe UI'
$tr.Font.Size = 15
$tr.Font.Color.RGB = (New-Rgb 255 255 255)
$tb.TextFrame.WordWrap = -1

$tb = $s.Shapes.AddTextbox(1, 520, 424, 370, 80)
$tr = $tb.TextFrame.TextRange
$tr.Text = "Проверил:`r_________________________"
$tr.Font.Name = 'Segoe UI'
$tr.Font.Size = 15
$tr.Font.Color.RGB = (New-Rgb 255 255 255)
$tb.TextFrame.WordWrap = -1

$tb = $s.Shapes.AddTextbox(1, 70, 486, 820, 26)
$tr = $tb.TextFrame.TextRange
$tr.Text = '2026 г.'
$tr.Font.Name = 'Segoe UI'
$tr.Font.Size = 13
$tr.Font.Color.RGB = (New-Rgb 170 195 230)

$slides[$slides.Count - 1].Notes = 'Тема курсового проекта — разработка информационной системы для магазина автозапчастей «Автомастерская». Работа выполнена на платформе .NET с использованием Avalonia UI и СУБД SQL Server.'

# ============================ 2. Содержание ============================
$s = New-Slide -Id 'toc' -Title 'Содержание доклада' -Subtitle 'Структура курсовой работы'
$items = @(
    @{ n = '1'; t = 'Введение: актуальность, цель и задачи работы'; d = 'Объект, предмет, методы проектирования' },
    @{ n = '2'; t = 'Теоретико-аналитическая часть'; d = 'Предметная область, анализ аналогов, выбор технологий, логическая модель данных' },
    @{ n = '3'; t = 'Проектно-программная часть'; d = 'База данных, архитектура приложения, алгоритмы, интерфейс, права доступа, тестирование' },
    @{ n = '4'; t = 'Заключение и практические результаты'; d = 'Достижение цели, значимость, направления развития' }
)
$y = 128
foreach ($it in $items) {
    $c = Add-Card -Slide $s -Left 60 -Top $y -Width 840 -Height 78
    $num = $s.Shapes.AddTextbox(1, 76, ($y + 16), 52, 46)
    $nt = $num.TextFrame.TextRange
    $nt.Text = $it.n
    $nt.Font.Name = 'Segoe UI Semibold'
    $nt.Font.Size = 30
    $nt.Font.Bold = -1
    $nt.Font.Color.RGB = (New-Rgb 197 90 17)
    $tb = $s.Shapes.AddTextbox(1, 138, ($y + 12), 740, 60)
    $tr = $tb.TextFrame.TextRange
    $tr.Text = $it.t
    $tr.Font.Name = 'Segoe UI Semibold'
    $tr.Font.Size = 17
    $tr.Font.Bold = -1
    $tr.Font.Color.RGB = (New-Rgb 31 56 100)
    $tb2 = $s.Shapes.AddTextbox(1, 138, ($y + 40), 740, 30)
    $t2 = $tb2.TextFrame.TextRange
    $t2.Text = $it.d
    $t2.Font.Name = 'Segoe UI'
    $t2.Font.Size = 13
    $t2.Font.Color.RGB = (New-Rgb 89 89 89)
    $tb2.TextFrame.WordWrap = -1
    $y += 88
}
$slides[$slides.Count - 1].Notes = 'Доклад построен по структуре записки: введение, две главы, заключение.'

# ============================ 3. Введение ============================
$s = New-Slide -Id 'intro' -Title 'Введение: актуальность и цель работы' -Subtitle 'Постановка задачи курсового проекта'
$c = Add-Card -Slide $s -Left 50 -Top 116 -Width 420 -Height 158
$tb = $s.Shapes.AddTextbox(1, 68, 128, 386, 138)
$tr = $tb.TextFrame.TextRange
$tr.Text = 'АКТУАЛЬНОСТЬ'
$tr.Font.Name = 'Segoe UI Semibold'
$tr.Font.Size = 13
$tr.Font.Bold = -1
$tr.Font.Color.RGB = (New-Rgb 197 90 17)
$tb = $s.Shapes.AddTextbox(1, 68, 152, 386, 112)
$tr = $tb.TextFrame.TextRange
$tr.Text = 'Данные о товарах, остатках и заказах хранятся в разных таблицах и на бумаге. Сотрудникам приходится повторно вводить сведения и вручную сверять остатки, что повышает риск ошибок.'
$tr.Font.Name = 'Segoe UI'
$tr.Font.Size = 13
$tr.Font.Color.RGB = (New-Rgb 38 38 38)
$tb.TextFrame.WordWrap = -1

$c = Add-Card -Slide $s -Left 490 -Top 116 -Width 420 -Height 158
$tb = $s.Shapes.AddTextbox(1, 508, 128, 386, 138)
$tr = $tb.TextFrame.TextRange
$tr.Text = 'ЦЕЛЬ РАБОТЫ'
$tr.Font.Name = 'Segoe UI Semibold'
$tr.Font.Size = 13
$tr.Font.Bold = -1
$tr.Font.Color.RGB = (New-Rgb 197 90 17)
$tb = $s.Shapes.AddTextbox(1, 508, 152, 386, 112)
$tr = $tb.TextFrame.TextRange
$tr.Text = 'Обеспечить централизованный учёт ассортимента и складских остатков, упростить подбор автозапчастей, автоматизировать оформление и отслеживание заказов, снизить вероятность ошибок.'
$tr.Font.Name = 'Segoe UI'
$tr.Font.Size = 13
$tr.Font.Color.RGB = (New-Rgb 38 38 38)
$tb.TextFrame.WordWrap = -1

$tb = $s.Shapes.AddTextbox(1, 50, 292, 860, 26)
$tr = $tb.TextFrame.TextRange
$tr.Text = 'РЕШАЕМЫЕ ЗАДАЧИ'
$tr.Font.Name = 'Segoe UI Semibold'
$tr.Font.Size = 13
$tr.Font.Bold = -1
$tr.Font.Color.RGB = (New-Rgb 197 90 17)

$tasks = @(
    'Проанализировать предметную область, выделить бизнес-процессы, бизнес-правила и роли участников',
    'Спроектировать и реализовать реляционную базу данных в СУБД SQL Server',
    'Разработать алгоритмы оформления и отмены заказа и реализовать их программными модулями',
    'Разработать пользовательский интерфейс и настроить разграничение прав доступа по ролям'
)
$bullets = @()
foreach ($t in $tasks) { $bullets += @{ Text = $t } }
[void](Add-Bullets -Slide $s -Items $bullets -Left 50 -Top 320 -Width 860 -Height 140 -FontSize 14)

$tb = $s.Shapes.AddTextbox(1, 50, 462, 860, 44)
$tr = $tb.TextFrame.TextRange
$tr.Text = 'Методы: анализ документации, функциональное моделирование, нормализация отношений, объектно-ориентированное проектирование, модульное тестирование.'
$tr.Font.Name = 'Segoe UI'
$tr.Font.Size = 12
$tr.Font.Italic = -1
$tr.Font.Color.RGB = (New-Rgb 89 89 89)
$tb.TextFrame.WordWrap = -1
$slides[$slides.Count - 1].Notes = 'Актуальность — разрозненный учёт данных и ручные сверки остатков. Четыре задачи: анализ предметной области, проектирование БД, алгоритмы, интерфейс и права доступа.'

# ============================ 4. Предметная область ============================
$s = New-Slide -Id 'domain' -Title 'Предметная область и участники процесса' -Subtitle 'Раздел 1.1 — бизнес-процессы магазина автозапчастей'
$tb = $s.Shapes.AddTextbox(1, 50, 116, 860, 60)
$tr = $tb.TextFrame.TextRange
$tr.Text = 'Основной процесс работы приложения: пользователь входит → выбирает товары → позиция переходит в корзину → подтверждает заказ → система оформляет заказ и фиксирует его в базе → товар выдаётся → заказ закрывается.'
$tr.Font.Name = 'Segoe UI'
$tr.Font.Size = 14
$tr.Font.Color.RGB = (New-Rgb 38 38 38)
$tb.TextFrame.WordWrap = -1

$rows = @(
    @('Участник', 'Выполняемые действия'),
    @('Гость', 'Просмотр каталога, выбор товаров, просмотр отзывов'),
    @('Покупатель', 'Оформление заказа, отзывы о товаре, история своих покупок'),
    @('Менеджер', 'Обработка и сопровождение заказов, смена статусов'),
    @('Администратор', 'Управление каталогом, пользователями и ролями, все заказы')
)
[void](Add-Table -Slide $s -Rows $rows -Left 50 -Top 190 -Width 860 -Height 130 -FontSize 13 -ColWidths @(220, 640))

$tb = $s.Shapes.AddTextbox(1, 50, 336, 860, 26)
$tr = $tb.TextFrame.TextRange
$tr.Text = 'КЛЮЧЕВЫЕ БИЗНЕС-ПРАВИЛА'
$tr.Font.Name = 'Segoe UI Semibold'
$tr.Font.Size = 13
$tr.Font.Bold = -1
$tr.Font.Color.RGB = (New-Rgb 197 90 17)

$rules = @(
    'Скидка 25 % — если в предыдущем календарном месяце у клиента не было позиций заказа',
    'Заказ считается крупным, если сумма позиций больше 5, иначе — малым',
    'Строка или заказ подсвечивается красным, если сумма не превышает 3',
    'Количество в позиции заказа положительное, остаток на складе после списания неотрицательный',
    'Удаление позиции возвращает зарезервированное количество запчастей на склад'
)
$b = @()
foreach ($r in $rules) { $b += @{ Text = $r } }
[void](Add-Bullets -Slide $s -Items $b -Left 50 -Top 364 -Width 860 -Height 140 -FontSize 13)
$slides[$slides.Count - 1].Notes = 'Четыре роли участников: Гость, Покупатель, Менеджер, Администратор. Ключевые бизнес-правила: скидка 25 %, признак крупного заказа, красная подсветка, контроль остатков.'

# ============================ 4a. Диаграмма вариантов использования ============================
$s = New-Slide -Id 'usecase' -Title 'Диаграмма вариантов использования' -Subtitle 'Раздел 1.1 — взаимодействие участников с системой'
[void](Add-PictureFit -Slide $s -Path (Join-Path $diag 'use_case.png') -Left 60 -Top 112 -MaxWidth 300 -MaxHeight 330)
[void](Add-Caption -Slide $s -Text 'Рисунок — Диаграмма вариантов использования системы' -Left 60 -Top 452 -Width 300)

$actors = @(
    @{ t = 'Неавторизованный пользователь'; d = 'Просмотр товаров, фильтрация по категориям, вход в систему' },
    @{ t = 'Авторизованный пользователь'; d = 'Выбор товаров, оформление заказа, отзывы, история покупок, отказ от заказа' },
    @{ t = 'Сотрудник'; d = 'Обработка заказа, смена статуса, формирование отчёта' },
    @{ t = 'Администратор'; d = 'Управление пользователями, базами данных магазина и составом каталога' }
)
$y = 112
foreach ($a in $actors) {
    $c = Add-Card -Slide $s -Left 400 -Top $y -Width 500 -Height 60
    $tb = $s.Shapes.AddTextbox(1, 418, ($y + 7), 466, 22)
    $tr = $tb.TextFrame.TextRange
    $tr.Text = $a.t
    $tr.Font.Name = 'Segoe UI Semibold'
    $tr.Font.Size = 13
    $tr.Font.Bold = -1
    $tr.Font.Color.RGB = (New-Rgb 31 56 100)
    $tb = $s.Shapes.AddTextbox(1, 418, ($y + 28), 466, 28)
    $tr = $tb.TextFrame.TextRange
    $tr.Text = $a.d
    $tr.Font.Name = 'Segoe UI'
    $tr.Font.Size = 11
    $tr.Font.Color.RGB = (New-Rgb 89 89 89)
    $tb.TextFrame.WordWrap = -1
    $y += 66
}
$b = @(
    @{ Text = 'Кнопка входа и фильтрация по категориям доступны любому посетителю магазина' },
    @{ Text = 'Оформление и отказ от заказа доступны только авторизованному пользователю' },
    @{ Text = 'Обработка заказа и управление пользователями закреплены за сотрудником и администратором' }
)
[void](Add-Bullets -Slide $s -Items $b -Left 400 -Top 382 -Width 500 -Height 108 -FontSize 11)
$slides[$slides.Count - 1].Notes = 'На диаграмме четыре актёра: неавторизованный пользователь, авторизованный пользователь, сотрудник и администратор. Кнопка входа и фильтрация доступны всем, оформление и отказ от заказа — только авторизованному пользователю, обработка заказа — сотруднику, управление пользователями и базами магазина — администратору.'

# ============================ 5. Анализ аналогов ============================
$s = New-Slide -Id 'analog' -Title 'Сравнительный анализ существующих решений' -Subtitle 'Раздел 1.2 — обоснование собственной разработки'
$rows = @(
    @('Критерий', '1С: Магазин автозапчастей', 'МойСклад', 'ABCP'),
    @('Тип решения', 'Локальное на «1С:Предприятие 8»', 'Облачное', 'Специализированное для интернет-магазинов'),
    @('Оформление заказов', 'Есть', 'Есть', 'Есть'),
    @('Скидка 25 % по правилу', 'Есть', 'Есть', 'Нет'),
    @('Оценка «много / мало»', 'Нет', 'Нет', 'Нет'),
    @('Работа без интернета', 'Есть', 'Нет', 'Нет'),
    @('Стоимость внедрения', 'От 39 100 руб.', 'От 1 400 руб./мес.', 'От 39 900 руб./год')
)
[void](Add-Table -Slide $s -Rows $rows -Left 50 -Top 120 -Width 860 -Height 250 -FontSize 12 -ColWidths @(230, 230, 200, 200))

$c = Add-Card -Slide $s -Left 50 -Top 390 -Width 860 -Height 92 -Fill (New-Rgb 255 244 230)
$tb = $s.Shapes.AddTextbox(1, 68, 404, 826, 72)
$tr = $tb.TextFrame.TextRange
$tr.Text = 'Вывод: ни одно из рассмотренных решений не поддерживает оценку заказа «много / мало» и не покрывает весь набор бизнес-правил предметной области. Разрабатываемое приложение восполняет этот пробел, сохраняя современный интерфейс и автономный режим работы.'
$tr.Font.Name = 'Segoe UI'
$tr.Font.Size = 13
$tr.Font.Color.RGB = (New-Rgb 38 38 38)
$tb.TextFrame.WordWrap = -1
$slides[$slides.Count - 1].Notes = 'Сравнивались 1С, МойСклад и ABCP. Ключевой вывод — ни один аналог не реализует оценку заказа «много/мало» и часть бизнес-правил.'

# ============================ 6. Выбор технологий ============================
$s = New-Slide -Id 'tech' -Title 'Обоснование выбора технологических средств' -Subtitle 'Раздел 1.4 — принятые проектные решения'
$rows = @(
    @('Задача', 'Принятое решение', 'Обоснование'),
    @('Хранение данных', 'SQL Server', 'Бесшовная интеграция с экосистемой Microsoft, корпоративная безопасность, встроенная аналитика'),
    @('Пользовательский интерфейс', 'Avalonia UI', 'Единая кодовая база под Windows, macOS, Linux, Android, iOS; XAML-разметка без привязки к ОС'),
    @('Доступ к данным', 'Entity Framework Core', 'ORM-фреймворк: работа с БД через объекты C#, а не прямой SQL'),
    @('Среда разработки', 'Visual Studio', 'IntelliSense, отладчик, конструктор интерфейсов, средства работы с БД'),
    @('Язык программирования', 'C#', 'Читаемый код, безопасная обработка данных, прямая интеграция с Avalonia и SQL Server'),
    @('Контроль версий', 'GitHub', 'Облачное хранение исходного кода и доступ к предыдущим версиям')
)
[void](Add-Table -Slide $s -Rows $rows -Left 50 -Top 120 -Width 860 -Height 280 -FontSize 12 -ColWidths @(190, 170, 500))
$c = Add-Card -Slide $s -Left 50 -Top 416 -Width 860 -Height 62
$tb = $s.Shapes.AddTextbox(1, 68, 428, 826, 44)
$tr = $tb.TextFrame.TextRange
$tr.Text = 'Решение в связке C# + Avalonia UI + Entity Framework Core + SQL Server обеспечивает кроссплатформенность, автономную работу и минимум повторного ввода данных.'
$tr.Font.Name = 'Segoe UI'
$tr.Font.Size = 13
$tr.Font.Color.RGB = (New-Rgb 38 38 38)
$tb.TextFrame.WordWrap = -1
$slides[$slides.Count - 1].Notes = 'Стек: SQL Server, Avalonia UI, Entity Framework Core, C#, Visual Studio, GitHub.'

# ============================ 7. Логическая модель ============================
$s = New-Slide -Id 'logical' -Title 'Проектирование логической модели данных' -Subtitle 'Раздел 1.5 — приведение отношений к третьей нормальной форме'
$rows = @(
    @('Форма', 'Требование', 'Принятое решение'),
    @('1НФ', 'Атрибуты атомарны, строки уникальны', 'Имя и фамилия разделены, дата заказа — в поле Date_Order, каждая позиция заказа — отдельная запись, суррогатные ключи'),
    @('2НФ', 'Неключевые атрибуты зависят от первичного ключа', 'Сведения о запчастях вынесены в Item, о категориях — в Category, о пользователях — в Users; в Orders — только дата и внешние ключи'),
    @('3НФ', 'Нет зависимости неключевых атрибутов друг от друга', 'Создана таблица Category (Id_Category, Name_Category), в Item хранится только Category_Id; роли вынесены в таблицу Roles')
)
[void](Add-Table -Slide $s -Rows $rows -Left 50 -Top 120 -Width 860 -Height 240 -FontSize 12 -ColWidths @(90, 250, 520))

$stats = @(
    @{ v = '7'; l = 'таблиц БД' },
    @{ v = '7'; l = 'внешних ключей' },
    @{ v = '43'; l = 'записи справочников' },
    @{ v = '3'; l = 'проекта решения' }
)
$x = 50
foreach ($st in $stats) {
    $c = Add-Card -Slide $s -Left $x -Top 380 -Width 205 -Height 86
    $tb = $s.Shapes.AddTextbox(1, $x, 392, 205, 44)
    $tr = $tb.TextFrame.TextRange
    $tr.Text = $st.v
    $tr.Font.Name = 'Segoe UI Semibold'
    $tr.Font.Size = 30
    $tr.Font.Bold = -1
    $tr.Font.Color.RGB = (New-Rgb 46 116 181)
    $tb.TextFrame.TextRange.ParagraphFormat.Alignment = 2
    $tb = $s.Shapes.AddTextbox(1, $x, 434, 205, 26)
    $tr = $tb.TextFrame.TextRange
    $tr.Text = $st.l
    $tr.Font.Name = 'Segoe UI'
    $tr.Font.Size = 12
    $tr.Font.Color.RGB = (New-Rgb 89 89 89)
    $tb.TextFrame.TextRange.ParagraphFormat.Alignment = 2
    $x += 220
}
$slides[$slides.Count - 1].Notes = 'Нормализация до 3НФ устраняет аномалии вставки, обновления и удаления. Итог — 7 таблиц, 7 внешних ключей, 43 записи.'

# ============================ 8. Физическая модель (ER) ============================
$s = New-Slide -Id 'er' -Title 'Физическая модель базы данных' -Subtitle 'Раздел 2.1 — семь таблиц на сервере SQL Server, база данных cursevork'
[void](Add-PictureFit -Slide $s -Path (Join-Path $diag 'er.png') -Left 60 -Top 116 -MaxWidth 840 -MaxHeight 330)
[void](Add-Caption -Slide $s -Text 'Рисунок — ER-диаграмма базы данных: Роли, Пользователи, Категории, Предметы, Заказы, Статусы заказов, Отзывы' -Left 60 -Top 456 -Width 840)
$slides[$slides.Count - 1].Notes = 'Схема из семи таблиц с семью внешними ключами. Типы данных и ограничения приведены в таблицах 7–13 записки.'

# ============================ 9. Архитектура ============================
$s = New-Slide -Id 'architecture' -Title 'Архитектура информационной системы' -Subtitle 'Раздел 2.2 — четыре уровня, каждый обращается только к следующему'
[void](Add-PictureFit -Slide $s -Path (Join-Path $diag 'architecture.png') -Left 190 -Top 112 -MaxWidth 580 -MaxHeight 300)
[void](Add-Caption -Slide $s -Text 'Рисунок — Архитектура ИС «Магазин автозапчастей»' -Left 60 -Top 420 -Width 840)

$layers = @(
    'Интерфейс пользователя — Avalonia (.axaml)',
    'Логика окон — модели представления (ViewModels)',
    'Бизнес-правила — сервисы (Services)',
    'Доступ к данным — контекст EF Core (Data)'
)
$x = 50
foreach ($l in $layers) {
    $c = Add-Card -Slide $s -Left $x -Top 448 -Width 205 -Height 44
    $tb = $s.Shapes.AddTextbox(1, ($x + 6), 456, 193, 32)
    $tr = $tb.TextFrame.TextRange
    $tr.Text = $l
    $tr.Font.Name = 'Segoe UI'
    $tr.Font.Size = 11
    $tr.Font.Color.RGB = (New-Rgb 38 38 38)
    $tb.TextFrame.WordWrap = -1
    $x += 220
}
$slides[$slides.Count - 1].Notes = 'Четыре уровня: UI, ViewModels, Services, Data. Разметка XAML не обращается к базе напрямую — связь появляется только при вызове метода сервиса.'

# ============================ 10. Слои / структура решения ============================
$s = New-Slide -Id 'layers' -Title 'Взаимодействие слоёв и структура решения' -Subtitle 'Разделы 2.2 — потоки вызовов и состав проектов'
[void](Add-PictureFit -Slide $s -Path (Join-Path $diag 'datastreams.png') -Left 50 -Top 116 -MaxWidth 420 -MaxHeight 200)
[void](Add-Caption -Slide $s -Text 'Схема взаимодействия слоёв приложения' -Left 50 -Top 322 -Width 420)
[void](Add-PictureFit -Slide $s -Path (Join-Path $diag 'solution_structure.png') -Left 490 -Top 116 -MaxWidth 420 -MaxHeight 200)
[void](Add-Caption -Slide $s -Text 'Структура решения и состав проектов' -Left 490 -Top 322 -Width 420)

$rows = @(
    @('Проект', 'Состав'),
    @('Ktpt.CourseOrder', 'Окна, модели представления, сервисы'),
    @('Ktpt.CourseOrder.Data', 'Сущности и доступ к данным'),
    @('Ktpt.CourseOrder.Auth', 'Авторизация и разграничение прав')
)
[void](Add-Table -Slide $s -Rows $rows -Left 50 -Top 358 -Width 860 -Height 116 -FontSize 12 -ColWidths @(250, 610))
$slides[$slides.Count - 1].Notes = 'Вправо по схеме — вызов операции и запрос, влево — данные, результат или сообщение об ошибке. Решение из трёх проектов.'

# ============================ 10a. Состав слоёв архитектуры ============================
$s = New-Slide -Id 'layersdetail' -Title 'Распределение классов по уровням архитектуры' -Subtitle 'Раздел 2.2 — представление, модель представления, сервисы, доступ к данным'
[void](Add-PictureFit -Slide $s -Path (Join-Path $diag 'layers.png') -Left 50 -Top 112 -MaxWidth 380 -MaxHeight 344)
[void](Add-Caption -Slide $s -Text 'Рисунок — Состав уровней архитектуры приложения' -Left 50 -Top 462 -Width 380)

$b = @(
    @{ Text = 'Представление (Views, XAML) — девять окон: каталог, карточка товара, корзина, отзывы, заказы, обработка заказов, администрирование, профиль, роли' },
    @{ Text = 'Модель представления (ViewModels) — команды интерфейса, поиск и фильтрация каталога, пересчёт суммы корзины, смена статусов' },
    @{ Text = 'Сервисы предметной области (Services) — OrderService, ItemService, CategoryService, ReviewService, CatalogAdminService, AuthService' },
    @{ Text = 'Доступ к данным (Data) — DbFactory, WorkshopDbContext, SchemaUpgrade, SeedData' },
    @{ Text = 'Единое хранилище — SQL Server (localdb)\Ander, база данных cursevork' }
)
[void](Add-Bullets -Slide $s -Items $b -Left 460 -Top 120 -Width 440 -Height 330 -FontSize 12)
$slides[$slides.Count - 1].Notes = 'Диаграмма показывает распределение классов по четырём уровням. Окна представления не обращаются к базе напрямую: запрос проходит через модель представления и сервис, а доступ к данным выполняет контекст EF Core.'

# ============================ 11. Модули ============================
$s = New-Slide -Id 'modules' -Title 'Назначение основных модулей приложения' -Subtitle 'Раздел 2.2 — таблица 14 записки'
$rows = @(
    @('Модуль', 'Назначение'),
    @('Views', 'Девять окон приложения и файлы разметки .axaml / .axaml.cs'),
    @('ViewModels', 'Логика представления и команды: MainViewModel, CartViewModel, OrdersViewModel и др.'),
    @('Services', 'Бизнес-правила: OrderService, ItemService, CategoryService, ReviewService, CatalogAdminService'),
    @('Data', 'Контекст WorkshopDbContext, сущности таблиц, фабрика DbFactory, наполнение SeedData'),
    @('Auth', 'Авторизация и разграничение прав: AuthService, Permissions, AppUser, AuthSession'),
    @('Models', 'Классы данных: Product, CartItem, AppServices и результат оформления заказа'),
    @('Converters', 'Преобразователи значений для интерфейса: BrushConverters, ImagePathConverter'),
    @('App.axaml', 'Тема, цветовые ресурсы и запуск приложения; строка подключения в Connection.txt')
)
[void](Add-Table -Slide $s -Rows $rows -Left 50 -Top 120 -Width 860 -Height 330 -FontSize 12 -ColWidths @(170, 690))
$slides[$slides.Count - 1].Notes = 'Восемь модулей: Views, ViewModels, Services, Data, Auth, Models, Converters, App.axaml.'

# ============================ 11a. Создание проекта и подключение зависимостей ============================
$s = New-Slide -Id 'devcreate' -Title 'Создание проекта и подключение зависимостей' -Subtitle 'Раздел 2.4 — каркас решения на Avalonia UI и Entity Framework Core'
[void](Add-PictureFit -Slide $s -Path (Join-Path $diag 'project_create.png') -Left 50 -Top 116 -MaxWidth 420 -MaxHeight 200)
[void](Add-Caption -Slide $s -Text 'Создание проекта Avalonia и файла решения в командной строке' -Left 50 -Top 322 -Width 420)
[void](Add-PictureFit -Slide $s -Path (Join-Path $diag 'package_efcore.png') -Left 490 -Top 116 -MaxWidth 420 -MaxHeight 200)
[void](Add-Caption -Slide $s -Text 'Подключение пакетов Entity Framework Core' -Left 490 -Top 322 -Width 420)

$b = @(
    @{ Text = 'dotnet new avalonia.app -n Ktpt.CourseOrder -f net10.0 — проект на шаблоне Avalonia XAML' },
    @{ Text = 'dotnet new sln -n Ktpt.CourseOrder — файл решения; dotnet add reference — ссылки между проектами' },
    @{ Text = 'dotnet add package Microsoft.EntityFrameworkCore.SqlServer — ORM-фреймворк и провайдер SQL Server версии 10.0.0' },
    @{ Text = 'Зависимости фиксируются в Ktpt.CourseOrder.csproj и восстанавливаются командой dotnet restore' }
)
[void](Add-Bullets -Slide $s -Items $b -Left 50 -Top 352 -Width 860 -Height 130 -FontSize 13)
$slides[$slides.Count - 1].Notes = 'Проект создан шаблоном Avalonia XAML, решение — командой dotnet new sln. Пакеты Entity Framework Core подключены через dotnet add package; версии фиксируются в файле проекта.'

# ============================ 12. Подключение к БД и сборка ============================
$s = New-Slide -Id 'dbconnect' -Title 'Подключение к базе данных и сборка проекта' -Subtitle 'Разделы 2.2 и 2.4 — фабрика DbFactory и команда dotnet build'
[void](Add-PictureFit -Slide $s -Path (Join-Path $diag 'code_dbfactory.png') -Left 50 -Top 116 -MaxWidth 420 -MaxHeight 210)
[void](Add-Caption -Slide $s -Text 'Фрагмент кода фабрики подключения к базе данных' -Left 50 -Top 332 -Width 420)
[void](Add-PictureFit -Slide $s -Path (Join-Path $diag 'build_output.png') -Left 490 -Top 116 -MaxWidth 420 -MaxHeight 210)
[void](Add-Caption -Slide $s -Text 'Успешная сборка решения командой dotnet build' -Left 490 -Top 332 -Width 420)

$b = @(
    @{ Text = 'Строка подключения хранится в файле Connection.txt рядом с исполняемым файлом' },
    @{ Text = 'Основной режим — SQL Server (localdb)\Ander, база данных cursevork' },
    @{ Text = 'При недоступности сервера фабрика автоматически переключается на локальный автономный режим autoshop.db' }
)
[void](Add-Bullets -Slide $s -Items $b -Left 50 -Top 362 -Width 860 -Height 110 -FontSize 13)
$slides[$slides.Count - 1].Notes = 'DbFactory переключается на автономный SQLite-режим, если SQL Server недоступен. Сборка решения проходит без ошибок и предупреждений.'

# ============================ 13. Алгоритм оформления ============================
$s = New-Slide -Id 'algcreate' -Title 'Алгоритм оформления заказа' -Subtitle 'Раздел 2.4 — метод OrderService.CreateOrderAsync'
[void](Add-PictureFit -Slide $s -Path (Join-Path $diag 'alg_create.png') -Left 300 -Top 112 -MaxWidth 360 -MaxHeight 320)
[void](Add-Caption -Slide $s -Text 'Схема алгоритма оформления заказа' -Left 300 -Top 440 -Width 360)

$b = @(
    @{ Text = 'Пустая корзина — выводится сообщение «Корзина пуста»' },
    @{ Text = 'Товары загружаются по идентификаторам из таблицы Item' },
    @{ Text = 'Проверяется достаточность остатка на складе' },
    @{ Text = 'Рассчитывается скидка 25 % постоянному покупателю' },
    @{ Text = 'Сумма свыше 5 единиц помечается как крупный заказ' },
    @{ Text = 'Оформление выполняется в транзакции, остатки уменьшаются' }
)
[void](Add-Bullets -Slide $s -Items $b -Left 50 -Top 128 -Width 235 -Height 330 -FontSize 12)
$tb = $s.Shapes.AddTextbox(1, 690, 140, 220, 60)
$tr = $tb.TextFrame.TextRange
$tr.Text = 'P = P₀ × (1 − d / 100)'
$tr.Font.Name = 'Consolas'
$tr.Font.Size = 17
$tr.Font.Bold = -1
$tr.Font.Color.RGB = (New-Rgb 197 90 17)
$tb.TextFrame.WordWrap = -1
$tb = $s.Shapes.AddTextbox(1, 690, 200, 220, 130)
$tr = $tb.TextFrame.TextRange
$tr.Text = "P — цена с учётом скидки, руб.;`rP₀ — базовая цена, руб.;`rd — размер скидки, %"
$tr.Font.Name = 'Segoe UI'
$tr.Font.Size = 12
$tr.Font.Color.RGB = (New-Rgb 89 89 89)
$tb.TextFrame.WordWrap = -1
$slides[$slides.Count - 1].Notes = 'Расчёт цены по формуле со скидкой 25 %. Метод CreateOrderAsync — строки 30–167 файла OrderService.cs.'

# ============================ 14. Алгоритм отмены ============================
$s = New-Slide -Id 'algcancel' -Title 'Алгоритм отмены заказа' -Subtitle 'Раздел 2.4 — метод OrderService.CancelOrderAsync'
[void](Add-PictureFit -Slide $s -Path (Join-Path $diag 'alg_cancel.png') -Left 300 -Top 112 -MaxWidth 360 -MaxHeight 320)
[void](Add-Caption -Slide $s -Text 'Схема алгоритма отмены заказа' -Left 300 -Top 440 -Width 360)

$b = @(
    @{ Text = 'В справочнике OrderStatus находится статус «Отменён»' },
    @{ Text = 'Заказ загружается вместе с товаром' },
    @{ Text = 'Если заказ уже отменён — возврат признака false, данные не меняются' },
    @{ Text = 'Статус заказа переводится в «Отменён»' },
    @{ Text = 'Зарезервированное количество возвращается на склад' }
)
[void](Add-Bullets -Slide $s -Items $b -Left 50 -Top 140 -Width 235 -Height 260 -FontSize 12)
$slides[$slides.Count - 1].Notes = 'Повторная отмена безопасно игнорируется. Товар возвращается на склад, что подтверждено тест-кейсом № 6.'

# ============================ 15. Интерфейс ============================
$s = New-Slide -Id 'ui' -Title 'Разработка пользовательского интерфейса' -Subtitle 'Раздел 2.5 — девять экранных форм в стиле Fluent'
[void](Add-PictureFit -Slide $s -Path (Join-Path $scr '01_main.png') -Left 50 -Top 116 -MaxWidth 420 -MaxHeight 210)
[void](Add-Caption -Slide $s -Text 'Главное окно каталога автозапчастей' -Left 50 -Top 332 -Width 420)
[void](Add-PictureFit -Slide $s -Path (Join-Path $scr '05_cart.png') -Left 490 -Top 116 -MaxWidth 420 -MaxHeight 210)
[void](Add-Caption -Slide $s -Text 'Корзина с выбранными позициями заказа' -Left 490 -Top 332 -Width 420)

$forms = @(
    'Главное окно каталога', 'Карточка товара', 'Отзывы о товаре', 'Корзина',
    'Мои заказы', 'Обработка заказов', 'Управление магазином', 'Профиль пользователя', 'Управление ролями'
)
$x = 50; $y = 362
foreach ($f in $forms) {
    $c = Add-Card -Slide $s -Left $x -Top $y -Width 176 -Height 34
    $tb = $s.Shapes.AddTextbox(1, ($x + 6), ($y + 8), 164, 22)
    $tr = $tb.TextFrame.TextRange
    $tr.Text = $f
    $tr.Font.Name = 'Segoe UI'
    $tr.Font.Size = 11
    $tr.Font.Color.RGB = (New-Rgb 38 38 38)
    $tr.ParagraphFormat.Alignment = 2
    $x += 190
    if ($x -gt 800) { $x = 50; $y += 42 }
}
$slides[$slides.Count - 1].Notes = 'Девять экранных форм. Переходы выполняются кнопками верхней панели главного окна, карточка товара открывается из карточки каталога.'

# ============================ 15a. Каталог и карточка товара ============================
$s = New-Slide -Id 'catalog' -Title 'Каталог автозапчастей и карточка товара' -Subtitle 'Раздел 2.5 — поиск, фильтрация и подробное описание позиции'
[void](Add-PictureFit -Slide $s -Path (Join-Path $scr '01_main.png') -Left 50 -Top 116 -MaxWidth 420 -MaxHeight 210)
[void](Add-Caption -Slide $s -Text 'Главное окно каталога автозапчастей' -Left 50 -Top 332 -Width 420)
[void](Add-PictureFit -Slide $s -Path (Join-Path $scr '02_product_details.png') -Left 490 -Top 116 -MaxWidth 420 -MaxHeight 210)
[void](Add-Caption -Slide $s -Text 'Карточка товара: описание, характеристики и отзывы' -Left 490 -Top 332 -Width 420)

$b = @(
    @{ Text = 'Каталог выводит позиции с изображением, названием, артикулом, ценой и остатком на складе' },
    @{ Text = 'Поиск по названию и фильтр по категориям выполняются через модель представления и сервис ItemService' },
    @{ Text = 'Карточка товара открывается из каталога и показывает описание, характеристики и отзывы покупателей' },
    @{ Text = 'Добавление позиции в корзину доступно авторизованному пользователю' }
)
[void](Add-Bullets -Slide $s -Items $b -Left 50 -Top 362 -Width 860 -Height 120 -FontSize 13)
$slides[$slides.Count - 1].Notes = 'Каталог и карточка товара. Список формируется запросом к таблице Item, поиск и фильтр выполняются на стороне модели представления, загрузка данных — через ItemService.'

# ============================ 15b. Корзина и отзывы ============================
$s = New-Slide -Id 'cartreviews' -Title 'Корзина покупателя и отзывы о товаре' -Subtitle 'Раздел 2.5 — состав заказа и обратная связь покупателей'
[void](Add-PictureFit -Slide $s -Path (Join-Path $scr '07_cart.png') -Left 50 -Top 116 -MaxWidth 420 -MaxHeight 210)
[void](Add-Caption -Slide $s -Text 'Корзина покупателя с выбранными позициями' -Left 50 -Top 332 -Width 420)
[void](Add-PictureFit -Slide $s -Path (Join-Path $scr '03_reviews.png') -Left 490 -Top 116 -MaxWidth 420 -MaxHeight 210)
[void](Add-Caption -Slide $s -Text 'Отзывы о товаре и форма добавления отзыва' -Left 490 -Top 332 -Width 420)

$b = @(
    @{ Text = 'Позиции корзины хранятся в CartItem и содержат количество, цену и итоговую сумму позиции' },
    @{ Text = 'Итоговая сумма пересчитывается при каждом изменении состава корзины' },
    @{ Text = 'Отзыв сохраняется с фиксацией автора и оценки, сервис ReviewService проверяет права пользователя' },
    @{ Text = 'Оформление заказа доступно только из корзины авторизованного покупателя' }
)
[void](Add-Bullets -Slide $s -Items $b -Left 50 -Top 362 -Width 860 -Height 120 -FontSize 13)
$slides[$slides.Count - 1].Notes = 'Корзина и отзывы. Состав корзины ведётся в модели представления, отзывы записываются через ReviewService с проверкой прав и фиксацией автора.'

# ============================ 15c. Мои заказы и профиль ============================
$s = New-Slide -Id 'buyerorders' -Title 'Заказы покупателя и профиль пользователя' -Subtitle 'Раздел 2.5 — история покупок и учётные данные'
[void](Add-PictureFit -Slide $s -Path (Join-Path $scr '04_my_orders.png') -Left 50 -Top 116 -MaxWidth 420 -MaxHeight 210)
[void](Add-Caption -Slide $s -Text 'Мои заказы: список заказов пользователя со статусами' -Left 50 -Top 332 -Width 420)
[void](Add-PictureFit -Slide $s -Path (Join-Path $scr '08_profile.png') -Left 490 -Top 116 -MaxWidth 420 -MaxHeight 210)
[void](Add-Caption -Slide $s -Text 'Профиль пользователя и смена пароля' -Left 490 -Top 332 -Width 420)

$b = @(
    @{ Text = 'Список «Мои заказы» формируется только по заказам текущего пользователя' },
    @{ Text = 'Для каждого заказа выводятся дата, состав, итоговая сумма и текущий статус' },
    @{ Text = 'Покупатель может отменить заказ, пока он не передан в обработку сотруднику' },
    @{ Text = 'В профиле доступны просмотр и изменение учётных данных и пароля' }
)
[void](Add-Bullets -Slide $s -Items $b -Left 50 -Top 362 -Width 860 -Height 120 -FontSize 13)
$slides[$slides.Count - 1].Notes = 'Окна «Мои заказы» и «Профиль». Заказы выбираются по идентификатору пользователя, отмена доступна через метод отмены заказа с возвратом остатка на склад.'

# ============================ 15d. Обработка заказов сотрудником ============================
$s = New-Slide -Id 'managerorders' -Title 'Обработка заказов сотрудником магазина' -Subtitle 'Раздел 2.5 — сопровождение заказа и смена статуса'
[void](Add-PictureFit -Slide $s -Path (Join-Path $scr '02_manager_orders.png') -Left 50 -Top 116 -MaxWidth 420 -MaxHeight 210)
[void](Add-Caption -Slide $s -Text 'Обработка заказов сотрудником магазина' -Left 50 -Top 332 -Width 420)
[void](Add-PictureFit -Slide $s -Path (Join-Path $scr '04_manager_orders.png') -Left 490 -Top 116 -MaxWidth 420 -MaxHeight 210)
[void](Add-Caption -Slide $s -Text 'Список заказов и смена статуса заказа' -Left 490 -Top 332 -Width 420)

$b = @(
    @{ Text = 'Сотрудник видит все заказы магазина независимо от пользователя, оформившего покупку' },
    @{ Text = 'Статус выбирается из справочника OrderStatus: оформление, обработка, выдача, отмена' },
    @{ Text = 'Изменение статуса сохраняется в таблице Orders и подтверждается на экране' },
    @{ Text = 'Отмена заказа возвращает зарезервированное количество запчастей на склад' }
)
[void](Add-Bullets -Slide $s -Items $b -Left 50 -Top 362 -Width 860 -Height 120 -FontSize 13)
$slides[$slides.Count - 1].Notes = 'Рабочее место сотрудника: список всех заказов, просмотр состава и смена статуса. Право на операции подтверждается модулем Auth.'

# ============================ 15e. Панель администратора ============================
$s = New-Slide -Id 'adminpanel' -Title 'Панель администратора магазина' -Subtitle 'Раздел 2.5 — управление каталогом, пользователями и заказами'
[void](Add-PictureFit -Slide $s -Path (Join-Path $scr '03_admin.png') -Left 50 -Top 116 -MaxWidth 420 -MaxHeight 210)
[void](Add-Caption -Slide $s -Text 'Панель администратора магазина' -Left 50 -Top 332 -Width 420)
[void](Add-PictureFit -Slide $s -Path (Join-Path $scr '05_admin.png') -Left 490 -Top 116 -MaxWidth 420 -MaxHeight 210)
[void](Add-Caption -Slide $s -Text 'Управление данными магазина: товары, категории, пользователи' -Left 490 -Top 332 -Width 420)

$b = @(
    @{ Text = 'Администратору доступны все экраны приложения, включая управление товарами и категориями' },
    @{ Text = 'Состав каталога изменяется через CatalogAdminService с проверкой обязательных полей' },
    @{ Text = 'Учётные записи пользователей и назначение ролей выполняются в панели администратора' },
    @{ Text = 'Все изменения сохраняются в базе данных cursevork и отображаются в каталоге сразу' }
)
[void](Add-Bullets -Slide $s -Items $b -Left 50 -Top 362 -Width 860 -Height 120 -FontSize 13)
$slides[$slides.Count - 1].Notes = 'Панель администратора. Здесь выполняется управление составом магазина, пользователями и заказами; сервис CatalogAdminService проверяет корректность вводимых данных.'

# ============================ 15f. Роли и гостевой режим ============================
$s = New-Slide -Id 'rolesui' -Title 'Управление ролями и гостевой режим' -Subtitle 'Раздел 2.6 — разграничение прав доступа на уровне интерфейса'
[void](Add-PictureFit -Slide $s -Path (Join-Path $scr '09_roles.png') -Left 50 -Top 116 -MaxWidth 420 -MaxHeight 210)
[void](Add-Caption -Slide $s -Text 'Управление ролями пользователей' -Left 50 -Top 332 -Width 420)
[void](Add-PictureFit -Slide $s -Path (Join-Path $scr '01_main_guest.png') -Left 490 -Top 116 -MaxWidth 420 -MaxHeight 210)
[void](Add-Caption -Slide $s -Text 'Каталог в режиме гостя: кнопки личных разделов скрыты' -Left 490 -Top 332 -Width 420)

$b = @(
    @{ Text = 'Роль пользователя определяется при входе и сохраняется в таблице Users' },
    @{ Text = 'Проверка операции выполняется методом Permissions.Allows перед выполнением действия' },
    @{ Text = 'В гостевом режиме скрыты кнопки «Мои заказы», «Заказы», «Управление» и «Роли»' },
    @{ Text = 'Элементы интерфейса скрываются по признаку роли, поэтому недоступные действия недоступны и в логике' }
)
[void](Add-Bullets -Slide $s -Items $b -Left 50 -Top 362 -Width 860 -Height 120 -FontSize 13)
$slides[$slides.Count - 1].Notes = 'Окно управления ролями и каталог в режиме гостя. Ограничения реализованы в модуле Auth: перечисление Permission перечисляет операции, Permissions.Allows проверяет допустимость действия для роли.'

# ============================ 16. Права доступа ============================
$s = New-Slide -Id 'roles' -Title 'Разграничение прав доступа' -Subtitle 'Раздел 2.6 — таблица 16 записки'
$rows = @(
    @('Операция', 'Гость', 'Покупатель', 'Менеджер', 'Администратор'),
    @('Просмотр каталога автозапчастей', 'Да', 'Да', 'Да', 'Да'),
    @('Оформление заказа', 'Нет', 'Да', 'Да', 'Да'),
    @('Просмотр своих заказов', 'Нет', 'Да', 'Да', 'Да'),
    @('Редактирование профиля', 'Нет', 'Да', 'Да', 'Да'),
    @('Добавление отзыва о товаре', 'Нет', 'Да', 'Да', 'Да'),
    @('Обработка и отмена заказов', 'Нет', 'Нет', 'Да', 'Да'),
    @('Управление товарами', 'Нет', 'Нет', 'Нет', 'Да'),
    @('Управление категориями', 'Нет', 'Нет', 'Нет', 'Да'),
    @('Управление пользователями и ролями', 'Нет', 'Нет', 'Нет', 'Да')
)
[void](Add-Table -Slide $s -Rows $rows -Left 50 -Top 120 -Width 860 -Height 280 -FontSize 12 -ColWidths @(356, 126, 126, 126, 126))
$c = Add-Card -Slide $s -Left 50 -Top 416 -Width 860 -Height 62
$tb = $s.Shapes.AddTextbox(1, 68, 428, 826, 44)
$tr = $tb.TextFrame.TextRange
$tr.Text = 'Права реализованы в модуле Ktpt.CourseOrder.Auth: перечисление Permission перечисляет доступные операции, а метод Permissions.Allows проверяет допустимость действия для роли пользователя.'
$tr.Font.Name = 'Segoe UI'
$tr.Font.Size = 13
$tr.Font.Color.RGB = (New-Rgb 38 38 38)
$tb.TextFrame.WordWrap = -1
$slides[$slides.Count - 1].Notes = 'Девять операций, четыре роли. Проверка доступа — через Permissions.Allows.'

# ============================ 17. Тестирование ============================
$s = New-Slide -Id 'testing' -Title 'Тестирование программного продукта' -Subtitle 'Раздел 2.7 — проверка бизнес-логики и сквозная проверка через UI Automation'
[void](Add-PictureFit -Slide $s -Path (Join-Path $diag 'testrun_output.png') -Left 500 -Top 116 -MaxWidth 410 -MaxHeight 220)
[void](Add-Caption -Slide $s -Text 'Результаты сквозной проверки покупки' -Left 500 -Top 342 -Width 410)

$rows = @(
    @('№', 'Проверяемый сценарий', 'Фактический результат'),
    @('1', 'Вход администратора (логин ander)', 'Вход выполнен, роль «Администратор»'),
    @('2', 'Вход с неверным паролем', 'Сообщение «Неверный логин или пароль», вход не выполнен'),
    @('3', 'Оплата корзины со скидкой', 'Создан заказ № 61 на 5 850,00 руб. (скидка 1 950,00 руб.)'),
    @('4', 'Остаток аккумулятора в таблице Item', 'Остаток уменьшился с 6 до 5'),
    @('5', 'Заказ в окне «Мои заказы»', 'Статус «Оформление заказа»'),
    @('6', 'Отмена заказа', 'Статус «Отменён», остаток восстановлен до 6'),
    @('7', 'Заказ при недостаточном остатке', 'Отказ в оформлении, запись не добавлена'),
    @('8', 'Заказ более чем из пяти позиций', 'Установлен признак IsLarge, выведено предупреждение'),
    @('9', 'Смена статуса заказа сотрудником', 'Статус изменён и сохранён в таблице Orders'),
    @('10', 'Работа без входа под ролью «Гость»', 'Кнопки «Мои заказы», «Заказы», «Управление», «Роли» скрыты')
)
[void](Add-Table -Slide $s -Rows $rows -Left 50 -Top 120 -Width 430 -Height 340 -FontSize 9 -ColWidths @(28, 175, 227))

$c = Add-Card -Slide $s -Left 500 -Top 366 -Width 410 -Height 94 -Fill (New-Rgb 232 244 232)
$tb = $s.Shapes.AddTextbox(1, 516, 378, 378, 76)
$tr = $tb.TextFrame.TextRange
$tr.Text = 'Все десять тест-кейсов завершились успешно: ожидаемые и фактические результаты совпали, база данных после прогонов осталась в корректном состоянии.'
$tr.Font.Name = 'Segoe UI'
$tr.Font.Size = 12
$tr.Font.Color.RGB = (New-Rgb 38 38 38)
$tb.TextFrame.WordWrap = -1
$slides[$slides.Count - 1].Notes = 'Десять тест-кейсов, все успешны. Сквозная проверка через UI Automation повторяет сценарий покупки целиком.'

# ============================ 18. Развёртывание ============================
$s = New-Slide -Id 'deploy' -Title 'Руководство пользователя и развёртывание' -Subtitle 'Раздел 2.8 — сценарии работы и порядок установки'
$tb = $s.Shapes.AddTextbox(1, 50, 116, 860, 26)
$tr = $tb.TextFrame.TextRange
$tr.Text = 'ОСНОВНЫЕ СЦЕНАРИИ РАБОТЫ'
$tr.Font.Name = 'Segoe UI Semibold'
$tr.Font.Size = 13
$tr.Font.Bold = -1
$tr.Font.Color.RGB = (New-Rgb 197 90 17)
$b = @(
    @{ Text = 'Гость — открывает каталог, применяет поиск и фильтр по категориям, просматривает отзывы' },
    @{ Text = 'Покупатель — добавляет товары в корзину, оформляет заказ, следит за статусом, при необходимости отменяет заказ' },
    @{ Text = 'Менеджер — обрабатывает заказы, подтверждает их, меняет статусы' },
    @{ Text = 'Администратор — управляет каталогом, пользователями и ролями' }
)
[void](Add-Bullets -Slide $s -Items $b -Left 50 -Top 144 -Width 860 -Height 130 -FontSize 13)

$tb = $s.Shapes.AddTextbox(1, 50, 286, 860, 26)
$tr = $tb.TextFrame.TextRange
$tr.Text = 'ПОРЯДОК РАЗВЁРТЫВАНИЯ'
$tr.Font.Name = 'Segoe UI Semibold'
$tr.Font.Size = 13
$tr.Font.Bold = -1
$tr.Font.Color.RGB = (New-Rgb 197 90 17)
$b = @(
    @{ Text = 'Создание локального экземпляра: sqllocaldb create "Ander"' },
    @{ Text = 'Создание базы данных cursevork и выполнение скриптов sql\SQLCreateTable.sql и sql\insert_sqlserver.sql' },
    @{ Text = 'Копирование файла Connection.txt со строкой подключения рядом с исполняемым файлом' },
    @{ Text = 'Запуск приложения: выполняется наполнение справочников SeedData, доступны демо-учётные записи' }
)
[void](Add-Bullets -Slide $s -Items $b -Left 50 -Top 314 -Width 860 -Height 150 -FontSize 13)
$slides[$slides.Count - 1].Notes = 'Развёртывание выполняется штатными средствами .NET. Скрипты формируют семь таблиц и наполняют их справочниками.'

# ============================ 19. Заключение ============================
$s = New-Slide -Id 'conclusion' -Title 'Заключение: достигнутые результаты' -Subtitle 'Основные итоги курсового проекта'
$res = @(
    @{ v = '1'; t = 'Проанализирована предметная область'; d = 'Выделены бизнес-процессы, сформулированы бизнес-правила, определены 4 роли участников' },
    @{ v = '2'; t = 'Спроектирована и реализована база данных'; d = '7 таблиц, 7 внешних ключей, 43 записи справочников в СУБД SQL Server' },
    @{ v = '3'; t = 'Реализованы алгоритмы и модули'; d = 'OrderService (оформление и отмена заказа), DbFactory, Permissions — 3 проекта решения' },
    @{ v = '4'; t = 'Разработан интерфейс и разграничение прав'; d = '9 экранных форм, 4 роли, 9 операций, проверка через UI Automation' }
)
$y = 118
foreach ($r in $res) {
    $c = Add-Card -Slide $s -Left 50 -Top $y -Width 860 -Height 62
    $nb = $s.Shapes.AddShape(9, 66, ($y + 14), 34, 34)   # msoShapeOval
    $nb.Fill.ForeColor.RGB = (New-Rgb 46 116 181)
    $nb.Line.Visible = 0
    $nb.Shadow.Visible = 0
    $nt = $nb.TextFrame.TextRange
    $nt.Text = $r.v
    $nt.Font.Name = 'Segoe UI Semibold'
    $nt.Font.Size = 15
    $nt.Font.Bold = -1
    $nt.Font.Color.RGB = (New-Rgb 255 255 255)
    $nt.ParagraphFormat.Alignment = 2
    $tb = $s.Shapes.AddTextbox(1, 116, ($y + 9), 780, 22)
    $tr = $tb.TextFrame.TextRange
    $tr.Text = $r.t
    $tr.Font.Name = 'Segoe UI Semibold'
    $tr.Font.Size = 15
    $tr.Font.Bold = -1
    $tr.Font.Color.RGB = (New-Rgb 31 56 100)
    $tb = $s.Shapes.AddTextbox(1, 116, ($y + 32), 780, 24)
    $tr = $tb.TextFrame.TextRange
    $tr.Text = $r.d
    $tr.Font.Name = 'Segoe UI'
    $tr.Font.Size = 12
    $tr.Font.Color.RGB = (New-Rgb 89 89 89)
    $tb.TextFrame.WordWrap = -1
    $y += 70
}
$c = Add-Card -Slide $s -Left 50 -Top 400 -Width 860 -Height 88 -Fill (New-Rgb 255 244 230)
$tb = $s.Shapes.AddTextbox(1, 68, 412, 826, 70)
$tr = $tb.TextFrame.TextRange
$tr.Text = 'Практическая значимость: сокращение времени обработки заказа, исключение повторного ввода данных, снижение вероятности ошибок при учёте остатков. Направление развития — печатные формы заявки, договора и ведомости записи, резервное копирование базы данных и перенос приложения в веб-версию.'
$tr.Font.Name = 'Segoe UI'
$tr.Font.Size = 12
$tr.Font.Color.RGB = (New-Rgb 38 38 38)
$tb.TextFrame.WordWrap = -1
$slides[$slides.Count - 1].Notes = 'Цель работы достигнута: разработана информационная система для магазина автозапчастей «Автомастерская».'

# ============================ 20. Источники ============================
$s = New-Slide -Id 'sources' -Title 'Список использованных источников' -Subtitle 'Основные источники — полный список из 24 позиций приведён в записке'
$sources = @(
    'ГОСТ Р 7.0.100–2018. Библиографическая запись. Библиографическое описание. Общие требования и правила составления.',
    'ГОСТ Р 7.0.5–2008. Библиографическая ссылка. Общие требования и правила составления.',
    'ГОСТ 2.105–2019. Единая система программной документации. Требования к оформлению программных документов.',
    'Документация Avalonia UI — официальный сайт фреймворка.',
    'Документация Entity Framework Core — официальный сайт Microsoft.',
    'Документация SQL Server — официальный сайт Microsoft.',
    'Документация Visual Studio — официальный сайт Microsoft.',
    'Документация .NET — официальный сайт Microsoft.',
    'Официальный сайт компании ООО «1С» — продукты и цены.',
    'Официальный сайт сервиса МойСклад.',
    'Официальный сайт платформы ABCP.'
)
$b = @()
foreach ($x2 in $sources) { $b += @{ Text = $x2 } }
[void](Add-Bullets -Slide $s -Items $b -Left 50 -Top 120 -Width 860 -Height 340 -FontSize 12)
$slides[$slides.Count - 1].Notes = 'Список оформлен по ГОСТ Р 7.0.100–2018: сначала нормативные документы, затем электронные источники. Дата обращения — 16.11.2026.'

# ============================ 21. Финал ============================
$s = Add-BlankSlide -Deck $Deck
[void]$slides.Add([pscustomobject]@{ Slide = $s; Id = 'end'; Title = 'Спасибо за внимание'; Notes = '' })
$bg = $s.Shapes.AddShape(1, 0, 0, 960, 540)
$bg.Fill.ForeColor.RGB = (New-Rgb 31 56 100)
$bg.Line.Visible = 0
$bg.Shadow.Visible = 0
$band = $s.Shapes.AddShape(1, 0, 268, 960, 4)
$band.Fill.ForeColor.RGB = (New-Rgb 197 90 17)
$band.Line.Visible = 0
$band.Shadow.Visible = 0
$tb = $s.Shapes.AddTextbox(1, 70, 200, 820, 70)
$tr = $tb.TextFrame.TextRange
$tr.Text = 'Спасибо за внимание!'
$tr.Font.Name = 'Segoe UI Semibold'
$tr.Font.Size = 40
$tr.Font.Bold = -1
$tr.Font.Color.RGB = (New-Rgb 255 255 255)
$tb = $s.Shapes.AddTextbox(1, 70, 292, 820, 60)
$tr = $tb.TextFrame.TextRange
$tr.Text = "Разработка информационной системы для магазина автозапчастей`rКурсовой проект · Копейкин Андрей · 2026 г."
$tr.Font.Name = 'Segoe UI'
$tr.Font.Size = 15
$tr.Font.Color.RGB = (New-Rgb 200 218 240)
$tb.TextFrame.WordWrap = -1
$slides[$slides.Count - 1].Notes = 'Готов ответить на вопросы.'

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
    $bak = Join-Path $root ('docs\backup_deck_' + (Get-Date -Format 'yyyy-MM-dd-HHmmss') + '.pptx')
    Copy-Item -LiteralPath $outPath -Destination $bak -Force
}
$Pres.SaveAs($outPath, 24)   # ppSaveAsOpenXMLPresentation
$Pres.Close()
$Deck.App.Quit()

Write-Output ('OK презентация сохранена: ' + $outPath)
Write-Output ('Слайдов: ' + $total)
