$ErrorActionPreference = 'Stop'
$env:SKILL_DIR = 'C:\Projects\Curse-work-N1\tools\docx'
$env:IN_DOCX   = 'C:\Users\Andrey\.minimax\v2\assets\2026\11\17\20-23-01-050-asset_20261117-202301-050_d65666492bd8_f7d74105-Копейкин_пояснительная_записка_курсового_проекта.docx'
$env:OUT_DOCX  = "$env:TEMP\docx-zapiska\zapiska-v3.docx"

. "$env:SKILL_DIR\DocxEdit.ps1"
Add-Type -AssemblyName System.IO.Compression.FileSystem

function Read-DocXml([string]$path) {
    $zip = [System.IO.Compression.ZipFile]::OpenRead($path)
    try {
        $e = $zip.GetEntry('word/document.xml')
        $sr = New-Object System.IO.StreamReader($e.Open(), [Text.Encoding]::UTF8)
        $t = $sr.ReadToEnd(); $sr.Close(); return $t
    } finally { $zip.Dispose() }
}

function Write-DocXml([string]$path, [string]$xml) {
    $tmp = "$path.tmp"
    Copy-Item -LiteralPath $path -Destination $tmp -Force
    $zip = [System.IO.Compression.ZipFile]::Open($tmp, 'Update')
    try {
        $e = $zip.GetEntry('word/document.xml')
        $s = $e.Open(); $s.SetLength(0)
        $sw = New-Object System.IO.StreamWriter($s, (New-Object System.Text.UTF8Encoding $false))
        $sw.Write($xml); $sw.Flush(); $sw.Close()
    } finally { $zip.Dispose() }
    Move-Item -LiteralPath $tmp -Destination $path -Force
}

# Insert a superscript marker right after the anchor phrase.
# The anchor may end in the middle of a text run, so that run is split into
#   [run with text-before][marker][run with text-after]
function Add-RefMarker {
    param([string]$Xml, [string]$AnchorText, [string]$Marker, [int]$Occurrence = 1)

    $hits = @()
    foreach ($m in $script:ParaRx.Matches($Xml)) {
        $t = Get-ParaText $m.Value
        if ($t.Contains($AnchorText)) { $hits += $m }
    }
    if ($hits.Count -lt $Occurrence) {
        Write-Output ("  ADD NOT FOUND: {0}" -f $AnchorText.Substring(0,[Math]::Min(46,$AnchorText.Length)))
        return $Xml
    }
    $m = $hits[$Occurrence - 1]
    $p = $m.Value
    $full = Get-ParaText $p
    $anchorStart = $full.IndexOf($AnchorText)
    $anchorEnd = $anchorStart + $AnchorText.Length

    $runs = [regex]::Matches($p, '(?s)<w:r(?:\s[^>]*)?>(?:(?!</w:r>).)*?</w:r>')
    $cum = ''
    $targetRun = $null
    $offsetInRun = 0
    $runStart = 0
    foreach ($r in $runs) {
        $tx = ''
        foreach ($tm in [regex]::Matches($r.Value, '<w:t[^>]*>([^<]*)</w:t>')) { $tx += $tm.Groups[1].Value }
        $runStart = $cum.Length
        $cum += $tx
        if ($cum.Length -ge $anchorEnd) {
            $targetRun = $r
            $offsetInRun = $anchorEnd - $runStart
            break
        }
    }
    if ($null -eq $targetRun) { return $Xml }

    $tMatch = [regex]::Match($targetRun.Value, '(?s)(<w:t(?:\s[^>]*)?>)(.*?)(</w:t>)')
    if (-not $tMatch.Success) { return $Xml }
    $runText = $tMatch.Groups[2].Value

    $markerRun = '<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:t xml:space="preserve"> ' + (Escape-XmlText $Marker) + '</w:t></w:r>'

    if ($offsetInRun -ge $runText.Length) {
        $pos = $targetRun.Index + $targetRun.Length
        $pNew = $p.Insert($pos, $markerRun)
    }
    else {
        $before = $runText.Substring(0, $offsetInRun)
        $after = $runText.Substring($offsetInRun)
        $runOpen = [regex]::Match($targetRun.Value, '^(?s)<w:r(?:\s[^>]*)?>').Value
        $rPr = [regex]::Match($targetRun.Value, '(?s)<w:rPr>.*?</w:rPr>').Value
        $newBefore = $runOpen + $rPr + '<w:t xml:space="preserve">' + (Escape-XmlText $before) + '</w:t></w:r>'
        $newAfter = $runOpen + $rPr + '<w:t xml:space="preserve">' + (Escape-XmlText $after) + '</w:t></w:r>'
        $replacement = $newBefore + $markerRun + $newAfter
        $pNew = $p.Remove($targetRun.Index, $targetRun.Length).Insert($targetRun.Index, $replacement)
    }

    return $Xml.Remove($m.Index, $m.Length).Insert($m.Index, $pNew)
}

Copy-Item -LiteralPath $env:IN_DOCX -Destination $env:OUT_DOCX -Force
$xml = Read-DocXml $env:OUT_DOCX

# ---------- STEP 1: extend the source list ----------
$LAST_SOURCE = 'Microsoft.EntityFrameworkCore.SqlServer: пакет NuGet'
$newSources = @(
    'Microsoft.Data.Sqlite: пакет NuGet: [сайт]. — URL: https://www.nuget.org/packages/Microsoft.Data.Sqlite (дата обращения: 17.11.2026).',
    'Microsoft .NET. Документация по .NET 10: [официальная документация]. — URL: https://learn.microsoft.com/ru-ru/dotnet/core/compatibility/ (дата обращения: 17.11.2026).',
    'Microsoft .NET. Создание проектов в командной строке: dotnet new: [официальная документация]. — URL: https://learn.microsoft.com/ru-ru/dotnet/core/tools/dotnet-new (дата обращения: 17.11.2026).',
    'Microsoft SQL Server. Типы данных (Transact-SQL): [официальная документация]. — URL: https://learn.microsoft.com/ru-ru/sql/t-sql/data-types/data-types-transact-sql (дата обращения: 17.11.2026).',
    'Microsoft SQL Server. Целочисленные типы int, bigint, smallint, tinyint (Transact-SQL): [официальная документация]. — URL: https://learn.microsoft.com/ru-ru/sql/t-sql/data-types/int-bigint-smallint-and-tinyint-transact-sql (дата обращения: 17.11.2026).',
    'Microsoft SQL Server. Транзакции (Transact-SQL): [официальная документация]. — URL: https://learn.microsoft.com/ru-ru/sql/t-sql/language-elements/transactions-transact-sql (дата обращения: 17.11.2026).',
    'Git. Система контроля версий: [сайт]. — URL: https://github.com/git/git (дата обращения: 17.11.2026).'
)
foreach ($s in $newSources) {
    $xml = Insert-ParaAfter -Xml $xml -AfterMatch $LAST_SOURCE -NewText $s -TemplateMatch $LAST_SOURCE -Occurrence 1
    $LAST_SOURCE = $s
}
Write-Output 'sources extended: 12 -> 19'

# ---------- STEP 2: in-text markers ----------
$marks = @(
    @{ a = 'Методы проектирования: анализ нормативной и технической документации'; m = '[1]' },
    @{ a = 'Процесс приведения отношений к третьей нормальной форме, устраняющий аномалии вставки'; m = '[1]' },
    @{ a = 'Результаты тестирования программного продукта — исходные данные, ожидаемый и фактический результат'; m = '[1]' },
    @{ a = 'Алгоритм оформления заказа реализован в методе OrderService.CreateOrderAsync'; m = '[2]' },
    @{ a = 'Алгоритм отмены заказа реализован в методе OrderService.CancelOrderAsync'; m = '[2]' },
    @{ a = 'Хранилище данных приложения — реляционная база данных cursevork на сервере SQL Server'; m = '[3]' },
    @{ a = 'Структура работы соответствует логике проектирования и включает введение, две главы'; m = '[4]' },
    @{ a = 'Тестирование выполнялось двумя способами: проверка бизнес-логики на уровне сервисов'; m = '[5]' },
    @{ a = 'Вывод: все запланированные функции реализованы и работают корректно'; m = '[5]' },
    @{ a = 'Приложение «Автозапчасти» разработано как графическое клиентское приложение на платформе .NET'; m = '[6]' },
    @{ a = 'Интерфейс приложения выполнен в светлой теме Fluent библиотеки Avalonia UI'; m = '[6]' },
    @{ a = 'Проект создан командой dotnet new avalonia.app -n Ktpt.CourseOrder'; m = '[7]' },
    @{ a = 'Библиотека Avalonia отвечает за создание окон, отображение элементов интерфейса'; m = '[8]' },
    @{ a = 'Для упрощения разработки интерфейса подключён пакет CommunityToolkit.Mvvm'; m = '[9]' },
    @{ a = 'В приложении применена слоистая архитектура'; m = '[10]' },
    @{ a = 'Для работы с базой данных подключены пакеты Microsoft.EntityFrameworkCore'; m = '[11]' },
    @{ a = 'Подключение к базе данных выполняет фабрика DbFactory'; m = '[12]' },
    @{ a = 'Если сервер недоступен, фабрика автоматически переключается на локальный файл autoshop.db'; m = '[12]' },
    @{ a = 'В таблицы загружено 4 роли, 5 пользователей, 6 категорий'; m = '[13]' },
    @{ a = 'Для работы требуются 64-разрядная Windows 10 или новее, .NET 10 Desktop Runtime'; m = '[14]' },
    @{ a = 'Проект создан командой dotnet new avalonia.app'; m = '[15]' },
    @{ a = 'Типы данных и ограничения приведены в таблицах 7–13'; m = '[15]' },
    @{ a = 'Уникальность строк обеспечивается суррогатными первичными ключами Id_User, Id_Category, Id_Item, Id_Order'; m = '[16]' },
    @{ a = 'Затем в транзакции уменьшаются остатки товаров и в таблицу Orders добавляются записи'; m = '[17]' },
    @{ a = 'Исходный код проекта размещён в репозитории GitHub'; m = '[18]' }
)

Write-Output ''
Write-Output '=== inserting markers ==='
foreach ($mk in $marks) {
    $before = $xml
    $xml = Add-RefMarker -Xml $xml -AnchorText $mk.a -Marker $mk.m
    if ($xml -eq $before) {
        Write-Output ("  FAILED {0} : {1}" -f $mk.m, $mk.a.Substring(0,[Math]::Min(40,$mk.a.Length)))
    } else {
        Write-Output ("  ok {0} : {1}" -f $mk.m, $mk.a.Substring(0,[Math]::Min(40,$mk.a.Length)))
    }
}

Write-DocXml $env:OUT_DOCX $xml
Write-Output ''
Write-Output ("written: {0}" -f $env:OUT_DOCX)
