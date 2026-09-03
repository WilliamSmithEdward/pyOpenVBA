# One identical payload under many extensions, so the compression
# decision can be attributed to the file type alone.
param([Parameter(Mandatory = $true)][string]$Target)

$ErrorActionPreference = 'Stop'
if (Test-Path $Target) { Remove-Item $Target -Force }

$engine = New-Object -ComObject DAO.DBEngine.120
$db = $engine.CreateDatabase($Target, ';LANGID=0x0409;CP=1252;COUNTRY=0')

$tdf = $db.CreateTableDef('Probe')
$f = $tdf.CreateField('Id', 4); $tdf.Fields.Append($f)
$f = $tdf.CreateField('Payload', 101); $tdf.Fields.Append($f)
$db.TableDefs.Append($tdf)

$exts = @(
    'txt', 'csv', 'xml', 'htm', 'html', 'rtf', 'log', 'ini', 'json', 'bas',
    'doc', 'docx', 'xls', 'xlsx', 'ppt', 'pptx', 'pdf', 'accdb',
    'bmp', 'gif', 'jpg', 'jpeg', 'jpe', 'jfif', 'png', 'tif', 'tiff', 'ico', 'emf', 'wmf',
    'zip', 'cab', 'gz', 'tgz', '7z', 'rar', 'lzh',
    'mp3', 'mp4', 'wav', 'avi', 'wmv', 'wma', 'mpg', 'mov',
    'exe', 'dll', 'dat', 'bin', 'iso', 'msi'
)
$dir = Join-Path $env:TEMP 'ext_probe'
if (Test-Path $dir) { Remove-Item $dir -Recurse -Force }
New-Item -ItemType Directory -Path $dir | Out-Null
# Compressible content, so a stored-raw result can only be the type.
$payload = [System.Text.Encoding]::ASCII.GetBytes(('A' * 512))

$rs = $db.OpenRecordset('Probe')
$i = 1
foreach ($ext in $exts) {
    $path = Join-Path $dir ("probe." + $ext)
    [System.IO.File]::WriteAllBytes($path, $payload)
    try {
        $rs.AddNew()
        $rs.Fields('Id').Value = $i
        $child = $rs.Fields('Payload').Value
        $child.AddNew(); $child.Fields('FileData').LoadFromFile($path); $child.Update()
        $child.Close()
        $rs.Update()
        $i = $i + 1
    } catch {
        try { $rs.CancelUpdate() } catch {}
        Write-Output ("REFUSED " + $ext)
    }
}
$rs.Close()
$db.Close()
Write-Output ("attached {0} extensions" -f $exts.Count)
