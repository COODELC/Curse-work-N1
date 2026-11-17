param(
    [Parameter(Mandatory=$true)][string]$SourceDocx,
    [Parameter(Mandatory=$true)][string]$OutDocx
)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem

$W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

$map = @{
    94  = 'Шаги процесса оформления заказа с указанием входных и выходных данных по каждому шагу приведены в таблице 1.'
    111 = 'Распределение операций между участниками процесса — гостем, покупателем, сотрудником и администратором — приведено в таблице 2.'
    120 = 'Сравнение существующих программных решений по критериям, определяющим выбор разрабатываемой системы, приведено в таблице 3.'
    134 = 'Принятые технологические решения и обоснование каждого из них приведены в таблице 5.'
    139 = 'Процесс приведения отношений к третьей нормальной форме, устраняющий аномалии вставки, удаления и обновления данных, приведён в таблице 6.'
    145 = 'Структура таблицы «Роли», предназначенной для хранения перечня ролей пользователей и разграничения прав доступа, приведена в таблице 7.'
    149 = 'Структура таблицы «Пользователи», предназначенной для хранения учётных записей пользователей, их ролей и контактных данных, приведена в таблице 8.'
    154 = 'Структура таблицы «Категории», предназначенной для классификации автозапчастей по товарным группам, приведена в таблице 9.'
    158 = 'Структура таблицы «Предметы», предназначенной для хранения каталога автозапчастей с их ценами, складскими остатками и характеристиками, приведена в таблице 10.'
    162 = 'Структура таблицы «Заказы», предназначенной для хранения позиций заказов покупателей, их сумм, скидок и текущего статуса, приведена в таблице 11.'
    166 = 'Структура таблицы «Статусы заказов», предназначенной для хранения справочника состояний, в которых может находиться заказ, приведена в таблице 12.'
    170 = 'Структура таблицы «Отзывы», предназначенной для хранения отзывов покупателей о приобретённых автозапчастях с оценкой и комментарием, приведена в таблице 13.'
    197 = 'Состав основных модулей приложения и назначение каждого из них приведены в таблице 14.'
    210 = 'Цветовая схема оформления интерфейса с указанием кода цвета и областей применения приведена в таблице 15.'
    245 = 'Матрица прав доступа пользователей в зависимости от их роли приведена в таблице 16.'
    253 = 'Результаты тестирования программного продукта — исходные данные, ожидаемый и фактический результат по каждому тест-кейсу — приведены в таблице 17 и на рисунке 14.'
}
$insertBeforeIndex = 128
$insertText = 'Функциональные требования к системе с указанием ожидаемого поведения приведены в таблице 4.'
$templateIndex = 134

function Set-ParagraphText($p, $ns, $text) {
    $ts = @($p.SelectNodes('.//w:t', $ns))
    if ($ts.Count -eq 0) { throw 'В абзаце нет w:t' }
    $ts[0].InnerText = $text
    for ($i = 1; $i -lt $ts.Count; $i++) { [void]$ts[$i].ParentNode.RemoveChild($ts[$i]) }
}

# --- 1. распаковка ---
$work = Join-Path $env:TEMP ('docx_repack_' + [Guid]::NewGuid().ToString('N'))
[System.IO.Compression.ZipFile]::ExtractToDirectory($SourceDocx, $work)
try {
    # --- 2. правка document.xml ---
    $docXmlPath = Join-Path $work 'word\document.xml'
    $xml = New-Object System.Xml.XmlDocument
    $xml.PreserveWhitespace = $true
    $xml.Load($docXmlPath)
    $ns = New-Object System.Xml.XmlNamespaceManager($xml.NameTable)
    $ns.AddNamespace('w', $W)
    $body = $xml.DocumentElement.SelectSingleNode('w:body', $ns)

    $paras = New-Object System.Collections.ArrayList
    foreach ($n in $body.ChildNodes) { if ($n.LocalName -eq 'p') { [void]$paras.Add($n) } }

    foreach ($k in ($map.Keys | Sort-Object)) {
        Set-ParagraphText $paras[[int]$k - 1] $ns $map[$k]
    }
    foreach ($k in ($map.Keys | Sort-Object)) {
        $p = $paras[[int]$k - 1]
        foreach ($hl in @($p.SelectNodes('.//w:highlight', $ns))) { [void]$hl.ParentNode.RemoveChild($hl) }
    }
    $target = $paras[$insertBeforeIndex - 1]
    $newP = $xml.ImportNode($paras[$templateIndex - 1], $true)
    Set-ParagraphText $newP $ns $insertText
    foreach ($hl in @($newP.SelectNodes('.//w:highlight', $ns))) { [void]$hl.ParentNode.RemoveChild($hl) }
    [void]$target.ParentNode.InsertBefore($newP, $target)

    $settings = New-Object System.Xml.XmlWriterSettings
    $settings.Encoding = New-Object System.Text.UTF8Encoding($false)
    $settings.Indent = $false
    $settings.OmitXmlDeclaration = $false
    $sw = [System.Xml.XmlWriter]::Create($docXmlPath, $settings)
    $xml.Save($sw); $sw.Close()

    # --- 3. упаковка вручную, канонический порядок записей ---
    $ordered = @()
    $ordered += '[Content_Types].xml'
    $ordered += '_rels/.rels'
    $all = Get-ChildItem -LiteralPath $work -Recurse -File | ForEach-Object {
        $_.FullName.Substring($work.Length + 1).Replace('\', '/')
    }
    $rest = $all | Where-Object { $ordered -notcontains $_ } | Sort-Object
    $ordered += $rest

    if (Test-Path -LiteralPath $OutDocx) { [System.IO.File]::Delete($OutDocx) }
    $out = [System.IO.Compression.ZipFile]::Open($OutDocx, 'Create')
    try {
        foreach ($rel in $ordered) {
            $src = Join-Path $work ($rel.Replace('/', '\'))
            $entry = $out.CreateEntry($rel, [System.IO.Compression.CompressionLevel]::Optimal)
            $es = $entry.Open()
            $bytes = [System.IO.File]::ReadAllBytes($src)
            $es.Write($bytes, 0, $bytes.Length)
            $es.Close()
        }
    } finally { $out.Dispose() }
    Write-Output ("Собрано: {0} (записей: {1}, {2} б)" -f $OutDocx, $ordered.Count, (Get-Item -LiteralPath $OutDocx).Length)
}
finally {
    try { Remove-Item -LiteralPath $work -Recurse -Force -ErrorAction SilentlyContinue } catch { }
}
