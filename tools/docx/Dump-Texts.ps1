param(
    [Parameter(Mandatory = $true)][string]$Docx,
    [int]$From = 0,
    [int]$To = 100000
)

Add-Type -AssemblyName System.IO.Compression.FileSystem
$ErrorActionPreference = 'Stop'

$zip = [System.IO.Compression.ZipFile]::OpenRead($Docx)
try {
    $entry = $zip.GetEntry('word/document.xml')
    $sr = New-Object System.IO.StreamReader($entry.Open(), [Text.Encoding]::UTF8)
    $xml = $sr.ReadToEnd()
    $sr.Close()
}
finally { $zip.Dispose() }

$paraRx = [regex]'(?s)<w:p\b(?:(?!</w:p>).)*?</w:p>|<w:p\b[^>]*/>'
$i = -1
foreach ($m in $paraRx.Matches($xml)) {
    $i++
    if ($i -lt $From -or $i -gt $To) { continue }
    $t = ''
    foreach ($tm in [regex]::Matches($m.Value, '<w:t[^>]*>([^<]*)</w:t>')) { $t += $tm.Groups[1].Value }
    $flags = ''
    if ($m.Value -match '<w:drawing') { $flags += ' DRAWING' }
    if ($m.Value -match '<w:object') { $flags += ' OBJECT' }
    if ($m.Value -match 'w:val="([^"]*)"') { }
    Write-Output ("[$i]$flags`t" + $t)
}
