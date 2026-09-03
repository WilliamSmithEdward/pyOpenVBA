# A second fixture: file types Access does not compress, and one large
# enough to spill out of a single row.
param([Parameter(Mandatory = $true)][string]$Target)

$ErrorActionPreference = 'Stop'
if (Test-Path $Target) { Remove-Item $Target -Force }

$engine = New-Object -ComObject DAO.DBEngine.120
$db = $engine.CreateDatabase($Target, ';LANGID=0x0409;CP=1252;COUNTRY=0')

$tdf = $db.CreateTableDef('Files')
$f = $tdf.CreateField('Id', 4); $tdf.Fields.Append($f)
$f = $tdf.CreateField('Payload', 101); $tdf.Fields.Append($f)
$db.TableDefs.Append($tdf)

# A zip (already compressed), a png, and 40 KB of text.
$zip = Join-Path $env:TEMP 'probe.zip'
$png = Join-Path $env:TEMP 'probe.png'
$big = Join-Path $env:TEMP 'probe_big.txt'
$src = Join-Path $env:TEMP 'zip_source'
if (Test-Path $src) { Remove-Item $src -Recurse -Force }
New-Item -ItemType Directory -Path $src | Out-Null
Set-Content -Path (Join-Path $src 'inner.txt') -Value 'inner' -Encoding ascii
if (Test-Path $zip) { Remove-Item $zip -Force }
Compress-Archive -Path (Join-Path $src '*') -DestinationPath $zip
# A minimal PNG: signature plus an IHDR-shaped chunk is enough for a
# by-extension decision.
$bytes = [byte[]](0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A) + [byte[]](1..64)
[System.IO.File]::WriteAllBytes($png, $bytes)
[System.IO.File]::WriteAllText($big, ('the quick brown fox ' * 2048))

$rs = $db.OpenRecordset('Files')
$i = 1
foreach ($path in $zip, $png, $big) {
    $rs.AddNew()
    $rs.Fields('Id').Value = $i
    $child = $rs.Fields('Payload').Value
    $child.AddNew(); $child.Fields('FileData').LoadFromFile($path); $child.Update()
    $child.Close()
    $rs.Update()
    $i = $i + 1
}
$rs.Close()
$db.Close()

foreach ($path in $zip, $png, $big) {
    $info = Get-Item $path
    Write-Output ("{0} {1} bytes" -f $info.Name, $info.Length)
}
Write-Output "built $Target"
