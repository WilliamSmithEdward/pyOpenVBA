# Ground truth for the Jet 3 (Access 97) reader.
#
#   dao_jet3.ps1 -Command build -Path db.mdb
#   dao_jet3.ps1 -Command dump  -Path db.mdb -Table AllTypes
#
# Access dropped Jet 3 in 2013, but Jet 4.0 -- which ships with Windows --
# still creates and reads it through DAO 3.6.  Both are 32-bit, so this
# script must run under SysWOW64 PowerShell; the gate arranges that.
#
# `dump` writes one JSON string: rows separated by LF, cells by TAB, each
# cell `Name=value` formatted as Format-Cell below, which mirrors the
# Python side in tests/test_live_access_jet3_gate.py.
#
# Dev-time oracle only; pyOpenVBA itself never touches COM.
param(
    [Parameter(Mandatory = $true)][string]$Command,
    [Parameter(Mandatory = $true)][string]$Path,
    [string]$Table = ""
)

$ErrorActionPreference = "Stop"
# Code page characters have to survive the pipe to a 64-bit Python.
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$inv = [System.Globalization.CultureInfo]::InvariantCulture
$dbVersion30 = 32
$dbFailOnError = 128

function Format-Cell($value) {
    if ($null -eq $value -or $value -is [System.DBNull]) { return "<null>" }
    if ($value -is [bool]) { if ($value) { return "True" } else { return "False" } }
    if ($value -is [decimal]) { return $value.ToString("0.0000", $inv) }
    if ($value -is [single] -or $value -is [double]) {
        $text = ([double]$value).ToString("0.######", $inv)
        if ($text -eq "-0") { return "0" }
        return $text
    }
    if ($value -is [datetime]) { return $value.ToString("yyyy-MM-dd HH:mm:ss", $inv) }
    if ($value -is [byte[]]) { return (($value | ForEach-Object { $_.ToString("x2") }) -join "") }
    if ($value -is [string]) { return $value.Replace("`t", "\t").Replace("`n", "\n") }
    return $value.ToString()
}

$dbe = New-Object -ComObject DAO.DBEngine.36

if ($Command -eq "build") {
    if (Test-Path $Path) { Remove-Item $Path -Force }
    $db = $dbe.CreateDatabase($Path, ";LANGID=0x0409;CP=1252;COUNTRY=0", $dbVersion30)
    # NOTE is a Jet reserved word, so every name here is bracketed.
    $db.Execute("CREATE TABLE AllTypes (Id COUNTER CONSTRAINT PK PRIMARY KEY, " +
                "Flag BIT, B BYTE, I SHORT, L LONG, Cur CURRENCY, Sng SINGLE, " +
                "Dbl DOUBLE, Dt DATETIME, T Text(80), M MEMO, O LONGBINARY, " +
                "Bin BINARY(12))", $dbFailOnError)
    $long = "long memo " * 400
    $rows = @(
        "(True, 200, -3000, 123456, 12.34, 1.5, 2.5, #2020-03-04 05:06:07#, 'text one', 'memo one')",
        "(False, 0, 32767, -2147483648, -0.5, -1.5, -2.5, #1899-12-30#, '', 'x')",
        "(True, 255, -32768, 2147483647, 922337203685477.5807, 3.4E38, 1.7E308, #9999-12-31 23:59:59#, 'accented: caf' & Chr(233) & ' na' & Chr(239) & 've', '$long')"
    )
    foreach ($values in $rows) {
        $db.Execute("INSERT INTO AllTypes (Flag, B, I, L, Cur, Sng, Dbl, Dt, T, M) VALUES $values", $dbFailOnError)
    }
    # A row with nothing but its key, and a deleted one, so the reader has
    # to handle an empty null mask and a dead slot.
    $db.Execute("INSERT INTO AllTypes (Flag) VALUES (False)", $dbFailOnError)
    $db.Execute("INSERT INTO AllTypes (Flag, T) VALUES (True, 'goes away')", $dbFailOnError)
    $db.Execute("DELETE FROM AllTypes WHERE T = 'goes away'", $dbFailOnError)

    # A second table with enough rows to spill across data pages, and an
    # index that is not the primary key.
    $db.Execute("CREATE TABLE Many (Id LONG CONSTRAINT PK2 PRIMARY KEY, [Note] Text(120))", $dbFailOnError)
    for ($i = 0; $i -lt 400; $i++) {
        $db.Execute("INSERT INTO Many (Id, [Note]) VALUES ($i, 'row $i " + ("pad " * 12) + "')", $dbFailOnError)
    }
    $db.Execute("CREATE INDEX ByNote ON Many ([Note])", $dbFailOnError)
    $db.Close()
    "ok"
    exit 0
}

if ($Command -eq "dump") {
    $db = $dbe.OpenDatabase($Path)
    $td = $db.TableDefs.Item($Table)
    $names = @($td.Fields | ForEach-Object { $_.Name })
    $rs = $db.OpenRecordset("SELECT * FROM [$Table] ORDER BY Id", 2)
    $lines = New-Object System.Collections.ArrayList
    while (-not $rs.EOF) {
        $cells = foreach ($n in $names) { "$n=" + (Format-Cell $rs.Fields.Item($n).Value) }
        [void]$lines.Add($cells -join "`t")
        $rs.MoveNext()
    }
    $rs.Close()
    $db.Close()
    ($lines -join "`n") | ConvertTo-Json -Compress
    exit 0
}

throw "unknown command $Command"
