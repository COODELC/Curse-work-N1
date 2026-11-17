$ErrorActionPreference = 'Stop'
$docPath = 'C:\Projects\Curse-work-N1\Копейкин _пояснительная_записка_курсового_проекта.docx'

$w = New-Object -ComObject Word.Application
$w.Visible = $false
$w.DisplayAlerts = 0
$w.Options.UpdateFieldsAtPrint = $false
$w.Options.Pagination = $true

$d = $w.Documents.Open($docPath, $false, $false)
Write-Output ('Открыт. Страниц до обновления: ' + $d.ComputeStatistics(2))

# Update all tables of contents (page numbers shift because 7 source
# paragraphs were added to the list).
foreach ($toc in $d.TablesOfContents) { $toc.Update() }
Write-Output 'Оглавление обновлено'

# Update the list of figures / tables if present.
foreach ($tof in $d.TablesOfFigures) { $tof.Update() }
Write-Output 'Перечень иллюстраций обновлён'

$d.Save()
Write-Output ('Страниц после: ' + $d.ComputeStatistics(2) +
              '; слов: ' + $d.ComputeStatistics(0) +
              '; таблиц: ' + $d.Tables.Count +
              '; рисунков: ' + $d.InlineShapes.Count)
$d.Close($false)
$w.Quit()
Write-Output 'OK сохранено'
