param([string]$DocxPath = 'F:\курсовая Копейкин Андрей\Копейкин Андрей\Копейкин Андрей Последняя версия — копия\Копейкин _пояснительная_записка_курсового_проекта.docx')

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem

$zip = [System.IO.Compression.ZipFile]::OpenRead((Resolve-Path $DocxPath).Path)
try {
    foreach ($part in @('word/document.xml', 'word/settings.xml')) {
        $entry = $zip.Entries | Where-Object { $_.FullName -eq $part }
        $reader = New-Object System.IO.StreamReader($entry.Open(), [System.Text.Encoding]::UTF8)
        $text = $reader.ReadToEnd()
        $reader.Close()

        Write-Output "########## $part ##########"

        foreach ($m in [regex]::Matches($text, '<w:instrText[^>]*>([^<]*)</w:instrText>')) {
            Write-Output ("INSTR: " + $m.Groups[1].Value.Trim())
        }

        if ($part -like '*settings.xml') {
            $u = [regex]::Match($text, '<w:updateFields[^>]*/>')
            Write-Output ("updateFields: " + $(if ($u.Success) { $u.Value } else { 'НЕТ' }))
        }
    }
}
finally { $zip.Dispose() }
