<#
.SYNOPSIS
    Наполнение пояснительной записки: заключение, список использованных
    источников, приложения А, Б, В и Г.
#>
param(
    [string]$Docx = 'Копейкин _пояснительная_записка_курсового_проекта.docx',
    [string]$ProjectRoot = '.'
)

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'DocxDom.ps1')

$W = $script:W
$docxPath = Join-Path $ProjectRoot $Docx
$backup = Join-Path (Join-Path $ProjectRoot 'docs') ('backup_zapiska_' + (Get-Date -Format 'yyyy-MM-dd-HHmmss') + '.docx')
Copy-Item -LiteralPath $docxPath -Destination $backup -Force
Write-Output ('Резервная копия: ' + $backup)

$screens = Join-Path $ProjectRoot 'docs\screens'
$code = Join-Path $ProjectRoot 'scr'

$ctx = New-DocxContext -Path $docxPath

# ---------- вспомогательные функции ----------

# Атрибут создаётся с явным префиксом w, иначе сериализатор .NET может
# выбрать другой префикс и Word посчитает документ повреждённым.
function Set-WAttr {
    param($El, [string]$Name, [string]$Value)
    $attr = $El.OwnerDocument.CreateAttribute('w', $Name, $script:W)
    $attr.Value = $Value
    $El.Attributes.Append($attr) | Out-Null
}

function New-Para {
    param(
        $Ctx,
        $Template,
        [string]$Text,
        [switch]$Code,
        [switch]$Heading
    )
    $p = $Template.CloneNode($true)
    $ns = New-Object System.Xml.XmlNamespaceManager($Ctx.Doc.NameTable)
    $ns.AddNamespace('w', $script:W)
    foreach ($d in @($p.SelectNodes('.//w:drawing', $ns))) { $d.ParentNode.RemoveChild($d) | Out-Null }
    foreach ($o in @($p.SelectNodes('.//w:object', $ns))) { $o.ParentNode.RemoveChild($o) | Out-Null }
    Set-ParaText -p $p -Text $Text | Out-Null

    if ($Code -or $Heading) {
        $doc = $Ctx.Doc
        $old = $p.SelectSingleNode('./w:pPr', $ns)
        $pPr = $doc.CreateElement('w', 'pPr', $script:W)

        if ($old) {
            $shd = $old.SelectSingleNode('./w:shd', $ns)
            if ($shd) { $pPr.AppendChild($shd.CloneNode($true)) | Out-Null }
        }

        $sp = $doc.CreateElement('w', 'spacing', $script:W)
        Set-WAttr -El $sp -Name 'before' -Value '0'
        Set-WAttr -El $sp -Name 'after' -Value '0'
        Set-WAttr -El $sp -Name 'line' -Value '240'
        Set-WAttr -El $sp -Name 'lineRule' -Value 'auto'
        $pPr.AppendChild($sp) | Out-Null

        $ind = $doc.CreateElement('w', 'ind', $script:W)
        Set-WAttr -El $ind -Name 'left' -Value '0'
        Set-WAttr -El $ind -Name 'firstLine' -Value '0'
        $pPr.AppendChild($ind) | Out-Null

        $jc = $doc.CreateElement('w', 'jc', $script:W)
        Set-WAttr -El $jc -Name 'val' -Value 'left'
        $pPr.AppendChild($jc) | Out-Null

        if ($Heading) {
            $kn = $doc.CreateElement('w', 'keepNext', $script:W)
            $pPr.InsertBefore($kn, $pPr.FirstChild) | Out-Null
        }

        if ($old) {
            $prPr = $old.SelectSingleNode('./w:rPr', $ns)
            if ($prPr) { $pPr.AppendChild($prPr.CloneNode($true)) | Out-Null }
        }

        if ($null -ne $old) { $p.ReplaceChild($pPr, $old) | Out-Null }
        else { $p.InsertBefore($pPr, $p.FirstChild) | Out-Null }
    }
    return $p
}

function Add-After {
    param($Ctx, $Node, $New)
    if ($New -is [System.Array]) {
        throw ('Add-After получил массив из ' + $New.Count + ' элементов вместо одного узла')
    }
    $Node.ParentNode.InsertAfter($New, $Node) | Out-Null
    return $New
}

# Абзац, начинающийся с указанного фрагмента (без учёта регистра).
function Find-ParaStartsWith {
    param($Ctx, [string]$Prefix)
    $blocks = Get-BodyBlocks $Ctx
    foreach ($p in $blocks) {
        if ($p.LocalName -ne 'p') { continue }
        $t = (Get-ParaTextOf $p).Trim()
        if ($t.StartsWith($Prefix)) { return $p }
    }
    return $null
}

try {
    Set-DocxScope -Ctx $ctx -AnchorMatch '2 ПРОЕКТНО-ПРОГРАММНАЯ ЧАСТЬ' | Out-Null

    # =========================================================
    # ЗАКЛЮЧЕНИЕ
    # =========================================================
    $head = Find-Para -Ctx $ctx -Match 'ЗАКЛЮЧЕНИЕ' -BodyLevel -Exact
    if ($null -eq $head) { throw 'Не найден заголовок ЗАКЛЮЧЕНИЕ' }

    $stubs = New-Object System.Collections.ArrayList
    $node = $head.NextSibling
    while ($null -ne $node) {
        if ($node.NodeType -ne 'Element') { $node = $node.NextSibling; continue }
        if ($node.LocalName -ne 'p') { break }
        $pPr = $node.SelectSingleNode('./w:pPr', $ctx.Nsm)
        $st = ''
        if ($pPr) { $s = $pPr.SelectSingleNode('w:pStyle', $ctx.Nsm); if ($s) { $st = $s.GetAttribute('val', $W) } }
        if ($st -eq '1') { break }
        [void]$stubs.Add($node)
        $node = $node.NextSibling
    }
    if ($stubs.Count -lt 3) { throw ('Заключение: найдено заглушек всего ' + $stubs.Count) }

    # шаблон обычного абзаца — второй заглушечный (без серой подсветки)
    $bodyTpl = $stubs[1].CloneNode($true)
    foreach ($s in $stubs) { $s.ParentNode.RemoveChild($s) | Out-Null }
    Write-Output ('Заключение: удалено заглушек — ' + $stubs.Count)

    $conclusion = @(
        'В ходе выполнения курсового проекта достигнута поставленная цель — разработана информационная система для магазина автозапчастей «Автомастерская», обеспечивающая централизованный учёт ассортимента и складских остатков, поиск и подбор автозапчастей, а также оформление и сопровождение заказов. Библиографическое описание использованных источников выполнено в соответствии с [5].',
        'При решении первой задачи выполнен анализ предметной области: выделены основные бизнес-процессы (учёт товаров и складских остатков, подбор автозапчастей, оформление и сопровождение заказов), сформулированы бизнес-правила и требования к системе, определены четыре роли участников — Гость, Покупатель, Менеджер и Администратор. Терминология предметной области приведена в соответствии с [3].',
        'При решении второй задачи спроектирована и реализована база данных в СУБД SQL Server [24]: разработана физическая модель из семи таблиц (Roles, Users, Category, Item, Orders, OrderStatus, ItemReview) с семью внешними ключами и нормализованными атрибутами, сопоставленными сущностям предметной области. Выполнено начальное наполнение базы данных — 43 записи справочников (4 роли, 6 категорий, 12 автозапчастей, 6 статусов заказа, 15 отзывов); учётные записи пользователей и демонстрационные данные заказов приложение создаёт при первом запуске.',
        'При решении третьей задачи разработаны алгоритмы оформления и отмены заказа и реализованы в виде программных модулей: сервис заказов OrderService (оформление заказа в транзакции с проверкой остатков, расчётом скидки 25 % постоянному покупателю и признаком крупного заказа свыше пяти единиц, отмена заказа с возвратом товара на склад), фабрика подключения к базе данных DbFactory (работа с SQL Server и автономный режим через SQLite [15, 19, 22, 23]) и модуль разграничения прав доступа Permissions. Решение разделено на три проекта — Ktpt.CourseOrder, Ktpt.CourseOrder.Data и Ktpt.CourseOrder.Auth, в составе приложения восемь модулей, состав которых приведён в таблице 14. Состав модулей соответствует стадиям создания автоматизированной системы [4], а управление зависимостями между ними реализовано по принципу внедрения зависимостей [16].',
        'При решении четвёртой задачи разработан пользовательский интерфейс на библиотеке Avalonia UI [7, 8] в стиле Fluent: реализовано девять экранных форм (главное окно каталога, карточка товара, отзывы о товаре, корзина, мои заказы, обработка заказов, управление магазином, профиль пользователя и управление ролями). Разграничение прав доступа настроено для четырёх ролей и девяти операций, логика представления отделена от представления по шаблону MVVM [11], оформление интерфейса задано стилями и ресурсами приложения [10].',
        'Проведено тестирование разработанного программного продукта: выполнено десять тест-кейсов, охватывающих проверку входа и разграничения прав, оформление и отмену заказа, контроль складских остатков, расчёт скидки, обработку недостаточного количества товара и сквозную проверку покупки. Требования к качеству программного продукта сформулированы по [6]; ожидаемые и фактические результаты всех тестов совпали, все запланированные функции реализованы и работают корректно.',
        'Практическая значимость результатов состоит в том, что разработанное приложение может быть использовано в магазине автозапчастей для сокращения времени обработки заказа, исключения повторного ввода данных, снижения вероятности ошибок при учёте остатков и контроля комплектации заказов. Развёртывание системы выполняется штатными средствами .NET [17].',
        'Направлением дальнейшего развития системы является расширение её функциональных возможностей: добавление печатных форм заявки, договора и ведомости записи, реализация резервного копирования базы данных и перенос приложения в веб-версию с использованием кэширования данных [18].'
    )
    $last = $head
    foreach ($t in $conclusion) {
        $np = New-Para -Ctx $ctx -Template $bodyTpl -Text $t
        $last = Add-After -Ctx $ctx -Node $last -New $np
    }
    Write-Output ('Заключение: добавлено абзацев — ' + $conclusion.Count)

    # =========================================================
    # СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ
    # =========================================================
    $srcHead = Find-Para -Ctx $ctx -Match 'СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ' -BodyLevel -Exact
    if ($null -eq $srcHead) { throw 'Не найден заголовок списка источников' }

    $gost1 = Find-Para -Ctx $ctx -Match 'ГОСТ Р 2.105-2019' -BodyLevel
    $gost2 = Find-Para -Ctx $ctx -Match 'ГОСТ 19.701-90' -BodyLevel
    if ($null -eq $gost1 -or $null -eq $gost2) { throw 'Не найдены записи ГОСТ в списке источников' }
    $srcTpl = $gost2.CloneNode($true)

    foreach ($m in @('20—25 источников', 'Фамилия, И. О. Название книги', 'Фамилия, И. О. Название статьи', 'Название ресурса : официальная', 'Далее ещё')) {
        $p = Find-ParaStartsWith -Ctx $ctx -Prefix $m
        if ($null -ne $p) {
            $p.ParentNode.RemoveChild($p) | Out-Null
            Write-Output ('Список источников: удалена заглушка — ' + $m)
        }
    }

    $sources = @(
        'ГОСТ 34.201-89. Автоматизированные системы. Термины и определения : межгосударственный стандарт. — Москва : Стандартинформ, 2010. — Текст : непосредственный.',
        'ГОСТ 34.601-90. Автоматизированные системы. Стадии создания автоматизированных систем. Требования к содержанию документов : межгосударственный стандарт. — Москва : Стандартинформ, 2010. — Текст : непосредственный.',
        'ГОСТ Р 7.0.100-2018. Библиографическая запись. Библиографическое описание. Общие требования и правила составления : национальный стандарт Российской Федерации. — Москва : Стандартинформ, 2018. — Текст : непосредственный.',
        'ГОСТ Р ИСО/МЭК 25010-2015. Системная и программная инженерия. Требования к оценке качества программных продуктов : национальный стандарт Российской Федерации. — Москва : Стандартинформ, 2015. — Текст : непосредственный.',
        'Avalonia : пакет NuGet. — Текст : электронный. — URL: https://www.nuget.org/packages/Avalonia (дата обращения: 16.11.2026).',
        'Avalonia UI. Знакомство с Avalonia : официальный сайт. — Текст : электронный. — URL: https://docs.avaloniaui.net/docs/get-started/ (дата обращения: 16.11.2026).',
        'Avalonia UI. Введение в привязку данных : официальный сайт. — Текст : электронный. — URL: https://docs.avaloniaui.net/docs/data-binding/introduction-to-data-binding (дата обращения: 16.11.2026).',
        'Avalonia UI. Стили : официальный сайт. — Текст : электронный. — URL: https://docs.avaloniaui.net/docs/styling/styles (дата обращения: 16.11.2026).',
        'Avalonia UI. Шаблон проектирования MVVM : официальный сайт. — Текст : электронный. — URL: https://docs.avaloniaui.net/docs/fundamentals/the-mvvm-pattern (дата обращения: 16.11.2026).',
        'Avalonia UI. Элемент управления Button : официальный сайт. — Текст : электронный. — URL: https://docs.avaloniaui.net/docs/controls/button/ (дата обращения: 16.11.2026).',
        'AvaloniaUI. Исходный код фреймворка : репозиторий GitHub. — Текст : электронный. — URL: https://github.com/AvaloniaUI/Avalonia (дата обращения: 16.11.2026).',
        'CommunityToolkit.Mvvm : пакет NuGet. — Текст : электронный. — URL: https://www.nuget.org/packages/CommunityToolkit.Mvvm (дата обращения: 16.11.2026).',
        'Microsoft.Data.Sqlite : пакет NuGet. — Текст : электронный. — URL: https://www.nuget.org/packages/Microsoft.Data.Sqlite (дата обращения: 16.11.2026).',
        'Microsoft .NET. Внедрение зависимостей : официальная документация. — Текст : электронный. — URL: https://learn.microsoft.com/ru-ru/dotnet/core/extensions/dependency-injection (дата обращения: 16.11.2026).',
        'Microsoft .NET. Инструменты командной строки : официальная документация. — Текст : электронный. — URL: https://learn.microsoft.com/ru-ru/dotnet/core/tools/ (дата обращения: 16.11.2026).',
        'Microsoft .NET. Кэширование в памяти : официальная документация. — Текст : электронный. — URL: https://learn.microsoft.com/ru-ru/dotnet/core/extensions/caching (дата обращения: 16.11.2026).',
        'Microsoft .NET. Основы Entity Framework Core : официальная документация. — Текст : электронный. — URL: https://learn.microsoft.com/ru-ru/ef/core/ (дата обращения: 16.11.2026).',
        'Microsoft .NET. Связи «один ко многим» в Entity Framework Core : официальная документация. — Текст : электронный. — URL: https://learn.microsoft.com/ru-ru/ef/core/modeling/relationships/one-to-many (дата обращения: 16.11.2026).',
        'Microsoft .NET. Транзакции и сохранение данных в Entity Framework Core : официальная документация. — Текст : электронный. — URL: https://learn.microsoft.com/ru-ru/ef/core/saving/transactions (дата обращения: 16.11.2026).',
        'Microsoft Entity Framework Core : пакет NuGet. — Текст : электронный. — URL: https://www.nuget.org/packages/Microsoft.EntityFrameworkCore (дата обращения: 16.11.2026).',
        'Microsoft Entity Framework Core для SQL Server : пакет NuGet. — Текст : электронный. — URL: https://www.nuget.org/packages/Microsoft.EntityFrameworkCore.SqlServer (дата обращения: 16.11.2026).',
        'SQL Server. Документация по системе управления базами данных : официальный сайт. — Текст : электронный. — URL: https://learn.microsoft.com/ru-ru/sql/sql-server/ (дата обращения: 16.11.2026).'
    )
    $last = $gost2
    foreach ($t in $sources) {
        $np = New-Para -Ctx $ctx -Template $srcTpl -Text $t
        $last = Add-After -Ctx $ctx -Node $last -New $np
    }
    Write-Output ('Список источников: добавлено записей — ' + $sources.Count + ' (всего 24)')

    # =========================================================
    # ПРИЛОЖЕНИЕ А
    # =========================================================
    $aHead = Find-Para -Ctx $ctx -Match 'ПРИЛОЖЕНИЕ A' -BodyLevel
    if ($null -eq $aHead) { throw 'Не найден заголовок приложения А' }
    $aIntro = New-Para -Ctx $ctx -Template $bodyTpl -Text 'В приложении приведена ER-диаграмма базы данных, отражающая состав и связи таблиц проектируемой информационной системы. Диаграмма построена по ГОСТ 19.701-90 [2] и соответствует структуре, приведённой в приложении Б.'
    $aHead.ParentNode.InsertAfter($aIntro, $aHead) | Out-Null
    Write-Output 'Приложение А: добавлен вводный абзац'

    # =========================================================
    # ПРИЛОЖЕНИЕ Б
    # =========================================================
    $b1 = Find-ParaStartsWith -Ctx $ctx -Prefix 'В.1 Создание таблиц'
    if ($null -eq $b1) { throw 'Не найден подзаголовок «В.1 Создание таблиц»' }
    Set-ParaText -p $b1 -Text 'Б.1 Создание таблиц' | Out-Null
    Write-Output 'Приложение Б: исправлен номер подзаголовка (В.1 → Б.1)'

    $bIntro = New-Para -Ctx $ctx -Template $bodyTpl -Text 'В приложении приведены скрипты создания базы данных и наполнения её данными. Скрипты выполнены в среде SQL Server [24] и соответствуют структуре, показанной на ER-диаграмме (рисунок А.1). Текст приложения оформлен по требованиям [1].'
    $b1.ParentNode.InsertBefore($bIntro, $b1) | Out-Null
    Write-Output 'Приложение Б: добавлен вводный абзац'

    # =========================================================
    # ПРИЛОЖЕНИЕ В
    # =========================================================
    $vIntro = Find-ParaStartsWith -Ctx $ctx -Prefix 'Приводятся только ключевые модули'
    if ($null -eq $vIntro) { throw 'Не найден вводный абзац приложения В' }
    Set-ParaText -p $vIntro -Text 'В приложении приведены листинги ключевых модулей разработанного программного продукта: сервиса заказов (реализация алгоритмов оформления и отмены заказа), фабрики подключения к базе данных и модуля разграничения прав доступа. Схемы алгоритмов, приведённые в разделе 2, построены в соответствии с [2]; состав модулей приложения приведён в таблице 14. Полный код проекта в приложение не включается — он размещается в репозитории системы контроля версий.' | Out-Null

    $codeTpl = Find-ParaStartsWith -Ctx $ctx -Prefix 'public void MethodName()'
    if ($null -eq $codeTpl) { throw 'Не найден шаблон абзаца листинга' }
    $codeTplClone = $codeTpl.CloneNode($true)
    $subTpl = Find-ParaStartsWith -Ctx $ctx -Prefix 'Г.1 Наименование модуля'
    if ($null -eq $subTpl) { throw 'Не найден подзаголовок приложения В' }
    $subTplClone = $subTpl.CloneNode($true)

    # удаляем заголовочный подзаголовок и абзацы-заглушки листинга
    $stubTexts = @('Г.1 Наименование модуля', 'public void MethodName()', '{', '// ВАШ КОД', '}')
    $heads = New-Object System.Collections.ArrayList
    $node = $subTpl
    $guard = 0
    while ($null -ne $node -and $guard -lt 12) {
        $guard++
        if ($node.NodeType -eq 'Element' -and $node.LocalName -eq 'p') {
            $t = (Get-ParaTextOf $node).Trim()
            if ($stubTexts -contains $t) { [void]$heads.Add($node) } else { break }
        }
        else { break }
        $node = $node.NextSibling
    }
    foreach ($h in $heads) { $h.ParentNode.RemoveChild($h) | Out-Null }
    Write-Output ('Приложение В: удалено абзацев-заглушек — ' + $heads.Count)

    $orderSrc = Get-Content -LiteralPath (Join-Path $code 'Ktpt.CourseOrder\Services\OrderService.cs') -Encoding UTF8
    $dbSrc = Get-Content -LiteralPath (Join-Path $code 'Ktpt.CourseOrder.Data\DbFactory.cs') -Encoding UTF8
    $permSrc = Get-Content -LiteralPath (Join-Path $code 'Ktpt.CourseOrder.Auth\Permissions.cs') -Encoding UTF8

    $listing = @(
        @{ Title = 'В.1 Сервис заказов Ktpt.CourseOrder.Services.OrderService — оформление заказа (метод CreateOrderAsync)';
           Lines = @($orderSrc[29..166]) },
        @{ Title = 'В.2 Сервис заказов Ktpt.CourseOrder.Services.OrderService — отмена заказа (метод CancelOrderAsync)';
           Lines = @($orderSrc[230..264]) },
        @{ Title = 'В.3 Фабрика подключения к базе данных Ktpt.CourseOrder.Data.DbFactory — метод Create';
           Lines = @($dbSrc[50..66]) },
        @{ Title = 'В.4 Модуль разграничения прав доступа Ktpt.CourseOrder.Auth.Permissions — метод Allows';
           Lines = @($permSrc[15..48]) }
    )

    $anchor = $vIntro
    foreach ($l in $listing) {
        $hp = New-Para -Ctx $ctx -Template $subTplClone -Text $l.Title -Heading
        $anchor = Add-After -Ctx $ctx -Node $anchor -New $hp
        foreach ($line in $l.Lines) {
            $cp = New-Para -Ctx $ctx -Template $codeTplClone -Text $line -Code
            $anchor = Add-After -Ctx $ctx -Node $anchor -New $cp
        }
    }
    Write-Output ('Приложение В: добавлено листингов — ' + $listing.Count)

    # =========================================================
    # ПРИЛОЖЕНИЕ Г
    # =========================================================
    $gIntro = Find-ParaStartsWith -Ctx $ctx -Prefix 'Каждая экранная форма сопровождается подписью'
    if ($null -eq $gIntro) { throw 'Не найден вводный абзац приложения Г' }
    Set-ParaText -p $gIntro -Text 'В приложении приведены снимки экрана основных экранных форм разработанного продукта. Нумерация рисунков в приложении своя: Рисунок Г.1, Рисунок Г.2. На рисунке Г.1 показан каталог автозапчастей, на рисунке Г.2 — корзина с выбранными позициями заказа [8, 10].' | Out-Null

    Set-Figure -Ctx $ctx -CaptionMatch 'Рисунок Д.1' -FilePath (Join-Path $screens '01_main.png') -NewCaption 'Рисунок Г.1 — Главное окно каталога автозапчастей' -WidthIn 6.3 | Out-Null
    Set-Figure -Ctx $ctx -CaptionMatch 'Рисунок Д.2' -FilePath (Join-Path $screens '05_cart.png') -NewCaption 'Рисунок Г.2 — Корзина с выбранными позициями заказа' -WidthIn 6.3 | Out-Null
    Write-Output 'Приложение Г: подставлены рисунки Г.1 и Г.2'

    Save-DocxContext -Ctx $ctx
    Write-Output 'Сохранено.'
}
finally {
    Close-DocxContext -Ctx $ctx
}
