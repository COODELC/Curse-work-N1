<#
.SYNOPSIS
    Замена изображений рисунков в пояснительной записке на исправленные
    PNG (светлая тема, починенные схемы алгоритмов, архитектура без обрезки)
    и вставка отсутствующего рисунка 10.
#>
param(
    [string]$Docx = 'Копейкин _пояснительная_записка_курсового_проекта.docx',
    [string]$ProjectRoot = '.'
)

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'DocxDom.ps1')

$docxPath = Join-Path $ProjectRoot $Docx
$backup = Join-Path (Join-Path $ProjectRoot 'docs') ('backup_zapiska_' + (Get-Date -Format 'yyyy-MM-dd-HHmmss') + '.docx')
Copy-Item -LiteralPath $docxPath -Destination $backup -Force
Write-Output ("Резервная копия: " + $backup)

$diag = Join-Path $ProjectRoot 'docs\diagrams'

$ctx = New-DocxContext -Path $docxPath
try {
    Set-DocxScope -Ctx $ctx -AnchorMatch '2 ПРОЕКТНО-ПРОГРАММНАЯ ЧАСТЬ' | Out-Null

    # --- Рисунок 2: схема взаимодействия слоёв (светлая тема) ---
    Set-Figure -Ctx $ctx -CaptionMatch 'Рисунок 2 ' -FilePath (Join-Path $diag 'datastreams.png') -WidthIn 6.3 | Out-Null

    # --- Рисунок 3: создание проекта Avalonia (код, светлая тема) ---
    Set-Figure -Ctx $ctx -CaptionMatch 'Рисунок 3 ' -FilePath (Join-Path $diag 'project_create.png') -WidthIn 6.3 -CaptionHasDrawing | Out-Null

    # --- Рисунок 4: структура решения ---
    Set-Figure -Ctx $ctx -CaptionMatch 'Рисунок 4 ' -FilePath (Join-Path $diag 'solution_structure.png') -WidthIn 6.3 | Out-Null

    # --- Рисунок 5: успешная сборка (dotnet-вывод, светлая тема) ---
    Set-Figure -Ctx $ctx -CaptionMatch 'Рисунок 5 ' -FilePath (Join-Path $diag 'build_output.png') -WidthIn 6.3 | Out-Null

    # --- Рисунок 6: архитектура ИС (была обрезана) ---
    Set-Figure -Ctx $ctx -CaptionMatch 'Рисунок 6 ' -FilePath (Join-Path $diag 'architecture.png') -WidthIn 5.6 | Out-Null

    # --- Рисунок 7: фрагмент кода фабрики (код, светлая тема) ---
    Set-Figure -Ctx $ctx -CaptionMatch 'Рисунок 7 ' -FilePath (Join-Path $diag 'code_dbfactory.png') -WidthIn 6.3 | Out-Null

    # --- Рисунок 8 / 9: скриншоты интерфейса (остаются без изменений) ---

    # --- Рисунок 10: подключение пакетов EF Core (отсутствовал) ---
    $cap10 = Find-Para -Ctx $ctx -Match 'Рисунок 10 ' -BodyLevel
    if ($null -eq $cap10) { throw 'Не найдена подпись рисунка 10' }
    if ($cap10.SelectSingleNode('.//w:drawing', $ctx.Nsm)) {
        Write-Output 'Рисунок 10 уже содержит изображение — пропускаем вставку'
    }
    else {
        $tpl = Find-Para -Ctx $ctx -Match 'Рисунок 9 ' -BodyLevel
        if ($null -eq $tpl) { throw 'Не найден шаблон подписи (Рисунок 9)' }
        $imgPara = $tpl.CloneNode($true)
        foreach ($t in @($imgPara.SelectNodes('.//w:t', $ctx.Nsm))) { $t.ParentNode.RemoveChild($t) | Out-Null }
        foreach ($d in @($imgPara.SelectNodes('.//w:drawing', $ctx.Nsm))) { $d.ParentNode.RemoveChild($d) | Out-Null }
        foreach ($o in @($imgPara.SelectNodes('.//w:object', $ctx.Nsm))) { $o.ParentNode.RemoveChild($o) | Out-Null }

        $png = Join-Path $diag 'package_efcore.png'
        $size = Get-PngSize -FilePath $png
        $cx = [int64][Math]::Round(6.3 * 914400)
        $cy = [int64][Math]::Round($cx * $size.H / $size.W)
        $media = Add-MediaFile -Ctx $ctx -FilePath $png
        $ctx.MaxDocPrId = (Get-MaxDocPrId $ctx) + 1
        $drawingXml = New-DrawingXml -RelId $media.RelId -FileName $media.Media -DocPrId $ctx.MaxDocPrId -Cx $cx -Cy $cy
        $frag = $ctx.Doc.CreateDocumentFragment()
        $frag.InnerXml = '<w:r xmlns:w="' + $script:W + '">' + $drawingXml + '</w:r>'
        $imgPara.AppendChild($frag) | Out-Null
        $cap10.ParentNode.InsertBefore($imgPara, $cap10) | Out-Null
        Write-Output ('OK вставлен рисунок 10: ' + $media.Media + ' (' + $cx + 'x' + $cy + ' EMU)')
    }

    # --- Рисунок 11: схема алгоритма оформления заказа ---
    Set-Figure -Ctx $ctx -CaptionMatch 'Рисунок 11 ' -FilePath (Join-Path $diag 'alg_create.png') -WidthIn 6.3 | Out-Null

    # --- Рисунок 12: схема алгоритма отмены заказа ---
    Set-Figure -Ctx $ctx -CaptionMatch 'Рисунок 12 ' -FilePath (Join-Path $diag 'alg_cancel.png') -WidthIn 6.3 | Out-Null

    # --- Рисунок 13: форма обработки заказов (исправляем подпись) ---
    $cap13 = Find-Para -Ctx $ctx -Match 'Рисунок 13 ' -BodyLevel
    if ($null -ne $cap13) {
        Set-ParaText -p $cap13 -Text 'Рисунок 13 — Форма обработки заказов' | Out-Null
        Write-Output 'OK исправлена подпись рисунка 13'
    }

    # --- Рисунок 14: результаты сквозной проверки (dotnet-вывод, светлая тема) ---
    Set-Figure -Ctx $ctx -CaptionMatch 'Рисунок 14 ' -FilePath (Join-Path $diag 'testrun_output.png') -WidthIn 6.3 | Out-Null

    # --- Приложение А: ER-диаграмма базы данных ---
    Set-Figure -Ctx $ctx -CaptionMatch 'Рисунок A.1' -FilePath (Join-Path $diag 'er.png') -WidthIn 6.3 | Out-Null

    Save-DocxContext -Ctx $ctx
    Write-Output 'Сохранено.'
}
finally {
    Close-DocxContext -Ctx $ctx
}
