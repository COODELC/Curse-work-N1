param(
    [Parameter(Mandatory=$true)][string]$TextDocx,
    [Parameter(Mandatory=$true)][string]$MediaDocx,
    [Parameter(Mandatory=$true)][string]$OutDocx
)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem

$base = Join-Path $env:TEMP ('docx_rebuild_' + [Guid]::NewGuid().ToString('N'))
$txt  = Join-Path $env:TEMP ('docx_text_'    + [Guid]::NewGuid().ToString('N'))
[System.IO.Compression.ZipFile]::ExtractToDirectory($MediaDocx, $base)
[System.IO.Compression.ZipFile]::ExtractToDirectory($TextDocx, $txt)

try {
    # --- связи из целого пакета ---
    $relsBase = Get-Content -LiteralPath (Join-Path $base 'word\_rels\document.xml.rels') -Encoding UTF8 -Raw
    $targetOfRid = @{}
    [regex]::Matches($relsBase, '<Relationship\s[^>]*Id="(rId\d+)"[^>]*Target="(media/[^"]+)"') | ForEach-Object {
        $targetOfRid[$_.Groups[1].Value] = $_.Groups[2].Value
    }
    $ridOfTarget = @{}
    foreach ($k in $targetOfRid.Keys) { $ridOfTarget[$targetOfTarget = $targetOfRid[$k]] = $k }

    # --- последовательность рисунков в каждом документе ---
    function Get-BlipTargets($docXml) {
        $out = New-Object System.Collections.ArrayList
        foreach ($m in [regex]::Matches($docXml, '<a:blip[^>]*r:embed="(rId\d+)"')) { [void]$out.Add($m.Groups[1].Value) }
        return $out
    }
    $txtDocPath = Join-Path $txt 'word\document.xml'
    $txtDoc = Get-Content -LiteralPath $txtDocPath -Encoding UTF8 -Raw

    # rels из текстового файла — нужны только для разрешения его rId в target
    $relsTxt = Get-Content -LiteralPath (Join-Path $txt 'word\_rels\document.xml.rels') -Encoding UTF8 -Raw
    $txtTarget = @{}
    [regex]::Matches($relsTxt, '<Relationship\s[^>]*Id="(rId\d+)"[^>]*Target="(media/[^"]+)"') | ForEach-Object {
        $txtTarget[$_.Groups[1].Value] = $_.Groups[2].Value
    }

    $baseBlips = Get-BlipTargets (Get-Content -LiteralPath (Join-Path $base 'word\document.xml') -Encoding UTF8 -Raw)
    $txtBlips  = Get-BlipTargets $txtDoc

    if ($baseBlips.Count -ne $txtBlips.Count) { throw ("Число рисунков не совпадает: база {0}, текст {1}" -f $baseBlips.Count, $txtBlips.Count) }

    # маппинг: позиция -> target базы -> rId базы
    $remap = @{}
    for ($i = 0; $i -lt $baseBlips.Count; $i++) {
        $oldRid = $txtBlips[$i]
        $newTarget = $targetOfRid[$baseBlips[$i]]
        $newRid = $ridOfTarget[$newTarget]
        if (-not $newRid) { throw ("Нет rId для {0}" -f $newTarget) }
        if ($oldRid -ne $newRid) { $remap[$oldRid] = $newRid }
    }
    Write-Output ("Переприсваивается rId: " + (($remap.GetEnumerator() | ForEach-Object { $_.Key + '->' + $_.Value }) -join ', '))

    # --- применяем переприсваивание к document.xml ---
    $newDoc = $txtDoc
    foreach ($k in $remap.Keys) {
        $newDoc = $newDoc -replace ('r:embed="' + [regex]::Escape($k) + '"'), ('r:embed="' + $remap[$k] + '"')
    }
    # проверка: не осталось ли v:imagedata / иных ссылок на media
    $leftover = [regex]::Matches($newDoc, 'r:(?:id|embed)="(rId\d+)"') | ForEach-Object { $_.Groups[1].Value } | Sort-Object -Unique
    $bad = @($leftover | Where-Object { $targetOfRid.ContainsKey($_) -and $targetOfRid[$_] -like 'media/*' -and -not $ridOfTarget.ContainsKey($targetOfRid[$_]) })
    if ($bad.Count -gt 0) { throw ('Нераспознанные ссылки на media: ' + ($bad -join ',')) }

    [System.IO.File]::WriteAllText($txtDocPath, $newDoc, (New-Object System.Text.UTF8Encoding($false)))

    # --- собираем пакет: содержимое из текстового файла, media и rels из целого ---
    $ordered = @('[Content_Types].xml', '_rels/.rels')
    $all = Get-ChildItem -LiteralPath $txt -Recurse -File | ForEach-Object {
        $_.FullName.Substring($txt.Length + 1).Replace('\', '/')
    }
    $skip = @('word/media/')
    $rest = $all | Where-Object {
        $rel = $_
        -not ($rel -like 'word/media/*') -and $rel -ne '[Content_Types].xml' -and $rel -ne '_rels/.rels'
    } | Sort-Object
    $ordered += $rest
    $ordered += 'word/_rels/document.xml.rels'

    if (Test-Path -LiteralPath $OutDocx) { [System.IO.File]::Delete($OutDocx) }
    $out = [System.IO.Compression.ZipFile]::Open($OutDocx, 'Create')
    try {
        foreach ($rel in $ordered) {
            $src = Join-Path $txt ($rel.Replace('/', '\'))
            if ($rel -eq 'word/_rels/document.xml.rels') { $src = Join-Path $base 'word\_rels\document.xml.rels' }
            if ($rel -eq '[Content_Types].xml') { $src = Join-Path $base '[Content_Types].xml' }
            $entry = $out.CreateEntry($rel, [System.IO.Compression.CompressionLevel]::Optimal)
            $es = $entry.Open()
            $bytes = [System.IO.File]::ReadAllBytes($src)
            $es.Write($bytes, 0, $bytes.Length); $es.Close()
        }
        # media из целого пакета
        Get-ChildItem -LiteralPath (Join-Path $base 'word\media') -File | ForEach-Object {
            $entry = $out.CreateEntry('word/media/' + $_.Name, [System.IO.Compression.CompressionLevel]::Optimal)
            $es = $entry.Open()
            $bytes = [System.IO.File]::ReadAllBytes($_.FullName)
            $es.Write($bytes, 0, $bytes.Length); $es.Close()
        }
    } finally { $out.Dispose() }
    Write-Output ("Собрано: {0} ({1} б)" -f $OutDocx, (Get-Item -LiteralPath $OutDocx).Length)
}
finally {
    try { Remove-Item -LiteralPath $base -Recurse -Force -ErrorAction SilentlyContinue } catch { }
    try { Remove-Item -LiteralPath $txt  -Recurse -Force -ErrorAction SilentlyContinue } catch { }
}
