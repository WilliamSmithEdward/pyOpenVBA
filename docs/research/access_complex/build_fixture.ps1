# Build a database with an attachment column and a multi-valued column,
# through DAO, so the storage they use can be measured.
param([Parameter(Mandatory = $true)][string]$Target)

$ErrorActionPreference = 'Stop'

if (Test-Path $Target) { Remove-Item $Target -Force }

$engine = New-Object -ComObject DAO.DBEngine.120
$db = $engine.CreateDatabase($Target, ';LANGID=0x0409;CP=1252;COUNTRY=0')

$dbLong = 4
$dbText = 10
$dbAttachment = 101
$dbComplexText = 109

$tdf = $db.CreateTableDef('Things')
$f = $tdf.CreateField('Id', $dbLong); $tdf.Fields.Append($f)
$f = $tdf.CreateField('Name', $dbText, 50); $tdf.Fields.Append($f)
$f = $tdf.CreateField('Files', $dbAttachment); $tdf.Fields.Append($f)
$f = $tdf.CreateField('Tags', $dbComplexText); $tdf.Fields.Append($f)
$db.TableDefs.Append($tdf)

$idx = $tdf.CreateIndex('PrimaryKey')
$idx.Fields.Append($idx.CreateField('Id'))
$idx.Primary = $true
$tdf.Indexes.Append($idx)

# Two small files to attach.
$one = Join-Path $env:TEMP 'attach_one.txt'
$two = Join-Path $env:TEMP 'attach_two.txt'
Set-Content -Path $one -Value 'hello from one' -Encoding ascii
Set-Content -Path $two -Value 'two' -Encoding ascii

$rs = $db.OpenRecordset('Things')
$rs.AddNew()
$rs.Fields('Id').Value = 1
$rs.Fields('Name').Value = 'first'
$files = $rs.Fields('Files').Value
$files.AddNew(); $files.Fields('FileData').LoadFromFile($one); $files.Update()
$files.AddNew(); $files.Fields('FileData').LoadFromFile($two); $files.Update()
$files.Close()
$tags = $rs.Fields('Tags').Value
foreach ($t in 'alpha', 'beta', 'gamma') {
    $tags.AddNew(); $tags.Fields('Value').Value = $t; $tags.Update()
}
$tags.Close()
$rs.Update()

$rs.AddNew()
$rs.Fields('Id').Value = 2
$rs.Fields('Name').Value = 'second'
$files = $rs.Fields('Files').Value
$files.AddNew(); $files.Fields('FileData').LoadFromFile($two); $files.Update()
$files.Close()
$tags = $rs.Fields('Tags').Value
$tags.AddNew(); $tags.Fields('Value').Value = 'delta'; $tags.Update()
$tags.Close()
$rs.Update()

# A row with neither, to see what an empty complex value stores.
$rs.AddNew()
$rs.Fields('Id').Value = 3
$rs.Fields('Name').Value = 'third'
$rs.Update()
$rs.Close()

$db.Close()
Write-Output "built $Target"
