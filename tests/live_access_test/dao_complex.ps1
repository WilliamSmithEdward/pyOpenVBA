# Read a table's attachment and multi-valued columns through DAO and
# print them as JSON, so what pyOpenVBA wrote is judged by the engine.
param(
    [Parameter(Mandatory = $true)][string]$Database,
    [Parameter(Mandatory = $true)][string]$Table,
    [string]$AttachmentColumn = '',
    [string]$MultiValueColumn = ''
)

$ErrorActionPreference = 'Stop'
$engine = New-Object -ComObject DAO.DBEngine.120
$db = $engine.OpenDatabase($Database)
$scratch = Join-Path $env:TEMP ('dao_complex_' + [System.Guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $scratch | Out-Null

$rows = @()
$serial = 0
$rs = $db.OpenRecordset($Table)
while (-not $rs.EOF) {
    $files = @()
    if ($AttachmentColumn -ne '') {
        $child = $rs.Fields($AttachmentColumn).Value
        while (-not $child.EOF) {
            $serial = $serial + 1
            $name = $child.Fields('FileName').Value
            $path = Join-Path $scratch ([string]$serial + '_' + $name)
            $child.Fields('FileData').SaveToFile($path)
            $bytes = [System.IO.File]::ReadAllBytes($path)
            $files += [pscustomobject]@{
                name = $name
                type = $child.Fields('FileType').Value
                size = $bytes.Length
                hex  = [System.BitConverter]::ToString($bytes).Replace('-', '').ToLower()
            }
            $child.MoveNext()
        }
        $child.Close()
    }
    $tags = @()
    if ($MultiValueColumn -ne '') {
        $child = $rs.Fields($MultiValueColumn).Value
        while (-not $child.EOF) {
            $tags += $child.Fields('Value').Value
            $child.MoveNext()
        }
        $child.Close()
    }
    $rows += [pscustomobject]@{
        id    = $rs.Fields('Id').Value
        files = @($files)
        tags  = @($tags)
    }
    $rs.MoveNext()
}
$rs.Close()
$db.Close()
Remove-Item $scratch -Recurse -Force

ConvertTo-Json -InputObject @($rows) -Depth 6 -Compress
