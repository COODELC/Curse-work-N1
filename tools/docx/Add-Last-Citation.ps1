$ErrorActionPreference = 'Stop'
$env:SKILL_DIR = 'C:\Projects\Curse-work-N1\tools\docx'
$env:IN_DOCX   = "$env:TEMP\docx-zapiska\zapiska-v3.docx"
$env:OUT_DOCX  = "$env:TEMP\docx-zapiska\zapiska-final.docx"

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

function Add-RefMarker {
    param([string]$Xml, [string]$AnchorText, [string]$Marker)

    $hits = @()
    foreach ($m in $script:ParaRx.Matches($Xml)) {
        $t = Get-ParaText $m.Value
        if ($t.Contains($AnchorText)) { $hits += $m }
    }
    if ($hits.Count -eq 0) {
        Write-Output ("  ADD NOT FOUND: {0}" -f $AnchorText.Substring(0,[Math]::Min(46,$AnchorText.Length)))
        return $Xml
    }
    $m = $hits[0]
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
        $pNew = $p.Remove($targetRun.Index, $targetRun.Length).Insert($targetRun.Index, $newBefore + $markerRun + $newAfter)
    }
    return $Xml.Remove($m.Index, $m.Length).Insert($m.Index, $pNew)
}

Copy-Item -LiteralPath $env:IN_DOCX -Destination $env:OUT_DOCX -Force
$xml = Read-DocXml $env:OUT_DOCX

# Source [19] = SQL Server documentation. It belongs where the DB engine
# and its schema capabilities are described.
$xml = Add-RefMarker -Xml $xml -AnchorText 'Схема состоит из семи таблиц: «Роли», «Пользователи», «Категории», «Предметы», «Заказы», «Статусы заказов» и «Отзывы»' -Marker '[19]'
Write-Output '  added [19] on the schema sentence'

Write-DocXml $env:OUT_DOCX $xml
Write-Output ("written: {0}" -f $env:OUT_DOCX)
