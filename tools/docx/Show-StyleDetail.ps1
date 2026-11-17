param([string]$DocxPath = 'F:\курсовая Копейкин Андрей\Копейкин Андрей\Копейкин Андрей Последняя версия — копия\Копейкин _пояснительная_записка_курсового_проекта.docx')

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem

$zip = [System.IO.Compression.ZipFile]::OpenRead((Resolve-Path $DocxPath).Path)
try {
    foreach ($part in @('word/styles.xml', 'word/numbering.xml')) {
        $entry = $zip.Entries | Where-Object { $_.FullName -eq $part }
        $reader = New-Object System.IO.StreamReader($entry.Open(), [System.Text.Encoding]::UTF8)
        $text = $reader.ReadToEnd()
        $reader.Close()

        Write-Output "########## $part ##########"

        if ($part -like '*styles.xml') {
            foreach ($id in @('a', '1', '2', '3', 'af0', 'af1', 'a4', 'Code', 'HTML0', 'my-2')) {
                $m = [regex]::Match($text, "<w:style [^>]*w:styleId=`"$id`"[^>]*>.*?</w:style>", 'Singleline')
                if ($m.Success) {
                    Write-Output ''
                    Write-Output "----- styleId=$id -----"
                    Write-Output (($m.Value -replace '><', ">`n<"))
                }
            }
        }
        else {
            # numId 6 и его abstractNum
            $m = [regex]::Match($text, '<w:num w:numId="6"[^>]*>.*?</w:num>', 'Singleline')
            if ($m.Success) {
                Write-Output '----- num numId=6 -----'
                Write-Output (($m.Value -replace '><', ">`n<"))

                $absId = [regex]::Match($m.Value, 'w:abstractNumId w:val="(\d+)"').Groups[1].Value
                $a = [regex]::Match($text, "<w:abstractNum w:abstractNumId=`"$absId`"[^>]*>.*?</w:abstractNum>", 'Singleline')

                if ($a.Success) {
                    Write-Output ''
                    Write-Output "----- abstractNum $absId (уровень 0) -----"
                    $lvl = [regex]::Match($a.Value, '<w:lvl w:ilvl="0"[^>]*>.*?</w:lvl>', 'Singleline')
                    Write-Output (($lvl.Value -replace '><', ">`n<"))
                }
            }
        }
    }
}
finally { $zip.Dispose() }
