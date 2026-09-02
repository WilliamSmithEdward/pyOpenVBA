# Ground truth for the storage engine: the ACE engine itself, driven through
# DAO (ACEDAO.DLL) with no Access window, no VBA and no dialogs.
#
#   dao_oracle.ps1 -Command build-alltypes -Path db.accdb -Rows 120
#   dao_oracle.ps1 -Command dump -Path db.accdb -Table AllTypes
#
# `dump` writes one JSON string to stdout: rows separated by LF, cells by
# TAB, each cell `Name=value` with values formatted as the Python side
# formats the engine's decoded values (see tests/test_live_access_engine_gate.py).
# Dev-time oracle only; pyOpenVBA itself never touches COM.
#
# Values go in as SQL literals wherever SQL has a literal, because
# PowerShell's COM binder cannot set Date, Currency, Decimal or BigInt
# fields ("Specified cast is not valid").  Binary columns, which have no
# literal, are set through the reflection binder, which marshals a byte[]
# as the SAFEARRAY DAO expects.
param(
    [Parameter(Mandatory = $true)][string]$Command,
    [Parameter(Mandatory = $true)][string]$Path,
    [int]$Rows = 120,
    [string]$Table = ""
)

$ErrorActionPreference = "Stop"
# Report where a COM call failed; the engine's own messages carry no location.
trap {
    $where = $_.InvocationInfo
    [Console]::Error.WriteLine("dao_oracle.ps1 line " + $where.ScriptLineNumber + ": " + $where.Line.Trim())
    [Console]::Error.WriteLine($_.Exception.Message)
    exit 1
}
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$inv = [System.Globalization.CultureInfo]::InvariantCulture
$dbFailOnError = 128
$dbOpenTable = 1
$dbOpenDynaset = 2
$dbOpenSnapshot = 4

function Invoke-Sql($target, [string]$sql, [int]$options) {
    try {
        if ($options -ge 0) { $target.Execute($sql, $options) } else { $target.Execute($sql) | Out-Null }
    } catch {
        throw ("SQL failed: " + $_.Exception.Message + "`n  " + $sql)
    }
}

function Set-Field($field, [byte[]]$value) {
    # PowerShell unrolls an array into separate arguments; build the
    # one-element argument list by hand so DAO receives one SAFEARRAY.
    $arguments = New-Object object[] 1
    $arguments[0] = $value
    [System.__ComObject].InvokeMember("Value", [System.Reflection.BindingFlags]::SetProperty, $null, $field, $arguments)
}

function Sql-Text([string]$s) { return "'" + $s.Replace("'", "''") + "'" }

function Pattern([int]$length, [int]$seed) {
    $out = New-Object byte[] $length
    for ($k = 0; $k -lt $length; $k++) { $out[$k] = [byte](($k * 31 + $seed) % 256) }
    return $out
}

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
    if ($value -is [string]) {
        # DAO renders a GUID field as "{guid {...}}".
        if ($value -match '^\{guid \{([0-9A-Fa-f-]{36})\}\}$') { return ("{" + $Matches[1] + "}").ToUpper() }
        return $value.Replace("`t", "\t").Replace("`n", "\n")
    }
    return $value.ToString()
}

# DAO's SQL dialect has no DECIMAL, so the table is created through the
# OLE DB provider (ANSI-92 SQL) and then filled through DAO.
function Create-AllTypes([string]$path) {
    $conn = New-Object -ComObject ADODB.Connection
    $conn.Open("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=$path")
    try {
        Invoke-Sql $conn ("CREATE TABLE AllTypes (Id AUTOINCREMENT PRIMARY KEY, Flag BIT, " +
            "Tiny BYTE, Small SHORT, Big LONG, Cash CURRENCY, Sgl SINGLE, Dbl DOUBLE, " +
            "Stamp DATETIME, Bin BINARY(50), Txt TEXT(100), Blob IMAGE, Story LONGTEXT, " +
            "Uid GUID, Frac DECIMAL(18,4), Huge BIGINT)") -1
    } finally {
        $conn.Close()
    }
}

# Every seventh row is all-null (a Yes/No column cannot be null, so it
# reads back False); every tenth Story is 5000 characters so long values
# chain across pages.
function Build-AllTypes($db, [int]$rowCount) {
    Invoke-Sql $db "CREATE INDEX IX_Txt ON AllTypes (Txt)" $dbFailOnError
    $cyrillic = [string]::Join("", [char]0x41F, [char]0x440, [char]0x438, [char]0x432, [char]0x435, [char]0x442)
    $cjk = [string]::Join("", [char]0x65E5, [char]0x672C)
    $cjk2 = [string]::Join("", [char]0x4E2D, [char]0x6587)
    for ($i = 1; $i -le $rowCount; $i++) {
        if ($i % 7 -eq 0) {
            Invoke-Sql $db "INSERT INTO AllTypes (Flag) VALUES (FALSE)" $dbFailOnError
            continue
        }
        $sign3 = if ($i % 3 -eq 0) { -1 } else { 1 }
        $sign5 = if ($i % 5 -eq 0) { -1 } else { 1 }
        $flag = if ($i % 2 -eq 0) { "TRUE" } else { "FALSE" }
        $tiny = $i % 256
        $small = (($i * 37) % 60000) - 30000
        $big = [int64]$i * 1000003 * $sign3
        $cash = ([double]$i / 8 * $sign5).ToString("0.####", $inv)
        $sgl = ([double]$i / 4).ToString("0.####", $inv)
        $dbl = ([double]$i * 1.5 + 0.25).ToString("0.####", $inv)
        $stamp = "#{0:D4}-{1:D2}-{2:D2} {3:D2}:{4:D2}:{5:D2}#" -f (1990 + ($i % 40)), (($i % 12) + 1), (($i % 28) + 1), ($i % 24), ($i % 60), (($i * 7) % 60)
        $txt = if ($i % 3 -eq 0) { "ascii only row $i" } else { "Row $i $cyrillic $cjk" }
        $storyLength = if ($i % 10 -eq 0) { 5000 } else { 20 }
        $story = (New-Object string ([char](65 + ($i % 26)), $storyLength)) + " end $i"
        if ($i % 4 -eq 0) { $story = $story + " " + $cjk2 }
        $uid = "{" + $i.ToString("X8") + "-1234-5678-9ABC-DEF012345678}"
        $frac = ([double]$i / 16 * $sign5).ToString("0.####", $inv)
        $huge = ([int64]$i * 10000000000 * $sign3).ToString($inv)
        Invoke-Sql $db ("INSERT INTO AllTypes (Flag, Tiny, Small, Big, Cash, Sgl, Dbl, Stamp, Txt, Story, Uid, Frac, Huge) VALUES (" +
            "$flag, $tiny, $small, $big, $cash, $sgl, $dbl, $stamp, " + (Sql-Text $txt) + ", " + (Sql-Text $story) + ", " +
            (Sql-Text $uid) + ", $frac, $huge)") $dbFailOnError
    }
    $rs = $db.OpenRecordset("SELECT Id, Bin, Blob FROM AllTypes ORDER BY Id", $dbOpenDynaset)
    while (-not $rs.EOF) {
        $i = [int]$rs.Fields.Item("Id").Value
        if ($i % 7 -ne 0) {
            $rs.Edit()
            Set-Field $rs.Fields.Item("Bin") ([byte[]](Pattern (($i % 50) + 1) $i))
            Set-Field $rs.Fields.Item("Blob") ([byte[]](Pattern ($i * 3) ($i + 1)))
            $rs.Update()
        }
        $rs.MoveNext()
    }
    $rs.Close()
    # One index per indexable column, so every key encoding is on disk,
    # plus a descending, a two-column and a unique one.
    foreach ($col in @("Flag", "Tiny", "Small", "Big", "Cash", "Sgl", "Dbl", "Stamp", "Bin", "Uid", "Frac", "Huge")) {
        Invoke-Sql $db "CREATE INDEX IX_$col ON AllTypes ($col)" $dbFailOnError
    }
    Invoke-Sql $db "CREATE INDEX IX_BigDesc ON AllTypes (Big DESC)" $dbFailOnError
    Invoke-Sql $db "CREATE INDEX IX_FlagTiny ON AllTypes (Flag, Tiny DESC)" $dbFailOnError
    Invoke-Sql $db "CREATE UNIQUE INDEX IX_UniqueBig ON AllTypes (Big) WITH IGNORE NULL" $dbFailOnError
}

# One row per BMP code point (surrogates aside) plus a few short strings,
# indexed, so the text sort keys the engine writes can be read back and
# turned into a collation table.  Characters the engine refuses are
# skipped; their code points are listed in the Skipped table.
function Build-Collation($db) {
    Invoke-Sql $db "CREATE TABLE Chars (Id AUTOINCREMENT PRIMARY KEY, Ch TEXT(255))" $dbFailOnError
    Invoke-Sql $db "CREATE INDEX IX_Ch ON Chars (Ch)" $dbFailOnError
    Invoke-Sql $db "CREATE TABLE Skipped (Cp LONG)" $dbFailOnError
    $rs = $db.OpenRecordset("Chars", $dbOpenTable)
    for ($cp = 1; $cp -le 0xFFFF; $cp++) {
        if ($cp -ge 0xD800 -and $cp -le 0xDFFF) { continue }
        try {
            $rs.AddNew()
            $rs.Fields.Item("Ch").Value = [string][char]$cp
            $rs.Update()
        } catch {
            try { $rs.CancelUpdate() } catch {}
            Invoke-Sql $db "INSERT INTO Skipped (Cp) VALUES ($cp)" $dbFailOnError
        }
    }
    # Composition rules: how extra weights and ignorable characters
    # combine and count positions.  Built from code points so this file
    # stays ASCII.
    $e = [string][char]0xE9      # e acute
    $E = [string][char]0xC9
    $sz = [string][char]0xDF     # sharp s, expands to ss
    $grave = [string][char]0x300 # combining grave
    $cyrP = [string][char]0x41F
    $cyrp = [string][char]0x43F
    $kanaA = [string][char]0x3042
    $cjk = [string][char]0x65E5
    $nbsp = [string][char]0xA0
    $ffi = [string][char]0xFB03          # ffi ligature, a three-byte primary
    $kanaSmallA = [string][char]0x3041
    $kanaGa = [string][char]0x304C       # voiced: a diacritic weight plus the kana suffix
    $kanaLong = [string][char]0x30FC
    $samples = @("aA", "Aa", "ab", "AB", "aa", "AA", "a a", "a  a", "a-a", "a_a", "a.a", "a'a", "ab-", "a-b-", "a--b",
                 "abc", "ABC", "aBc", "a1", "A1", "1a", "-a", "a-", "--", "---", " a", "a ", "  ", "-", "'",
                 "ee", "Ee", "eE", "EE", "eAb", "Eab", "ss", "SS", "ll", "LL", "ch", "CH", "I", "i",
                 $e, $E, ($e + "a"), ("a" + $e), ($e + $e), ("a" + $e + "a"), ($e + "a" + $e), ("ab" + $e), ($e + "ab"),
                 ("e" + $grave), ("e" + $grave + "a"), $sz, ($sz + "a"), ("a" + $sz), ($sz + "-"), ("s-"), ("ss-"),
                 $cyrP, $cyrp, ($cyrP + "a"), ("a" + $cyrP), ($cyrP + $cyrp),
                 $kanaA, ($kanaA + "a"), ("a" + $kanaA), ($kanaA + $kanaA), $cjk, ($cjk + "a"), ("a" + $cjk),
                 $nbsp, ("a" + $nbsp + "a"), ($e + "-"), ("-" + $e), ($e + "-" + $e),
                 ($cyrP + $E), ($E + $cyrP), ($cyrP + $cyrP + $E), ($cyrP + "-"), ($cyrP + $cyrP + "-"), ($cjk + "-"),
                 ($ffi + "-"), ($ffi + $E), ($kanaSmallA + $kanaA), ($kanaA + $kanaSmallA), ("a" + $kanaSmallA), ($kanaSmallA + "a"),
                 ($kanaGa + $kanaSmallA), ($kanaGa), ($kanaLong + $kanaA), ($E + $kanaA), ($kanaA + $E), ($kanaA + "-"), ("-" + $kanaA),
                 ($sz + $E), ("aa" + $E), ("a -"), ("- a"), ("a  "), ("a" * 255), (("b" * 100) + $E), (("b" * 254) + $E), ("c" * 254), ("c" * 255),
                 ($e + $grave), ("a" + $grave + $grave), ("A" + $E + $e), ($E + " " + $E))
    foreach ($s in $samples) {
        $rs.AddNew()
        $rs.Fields.Item("Ch").Value = $s
        $rs.Update()
    }
    # Element-count probes: each entry is a comma-separated list of code
    # points; a trailing hyphen's position code reveals how many
    # collation elements the characters before it produced.
    $probes = @("0xC6,0x2d", "0xDE,0x2d", "0xDF,0x2d", "0xE6,0x2d", "0xFE,0x2d", "0x132,0x2d", "0x133,0x2d", "0x152,0x2d", "0x153,0x2d", "0x1C4,0x2d", "0x1C5,0x2d", "0x1C6,0x2d", "0x1C7,0x2d", "0x1C8,0x2d", "0x1C9,0x2d", "0x1CA,0x2d", "0x1CB,0x2d", "0x1CC,0x2d", "0x1E2,0x2d", "0x1E3,0x2d", "0x1F1,0x2d", "0x1F2,0x2d", "0x1F3,0x2d", "0x1FC,0x2d", "0x1FD,0x2d", "0x5F0,0x2d", "0x5F1,0x2d", "0x5F2,0x2d", "0xFB00,0x2d", "0xFB01,0x2d", "0xFB02,0x2d", "0xFB03,0x2d", "0xFB04,0x2d", "0xFB05,0x2d", "0xFB06,0x2d", "0x3041,0x3041", "0x3042,0x3042,0x3041", "0x61,0x61,0x3041", "0x3041,0x3042,0x3041", "0x3041,0x2d", "0x3042,0x3042,0x3042,0x3041", "0x61,0x300", "0x65,0x301", "0x41,0x300", "0x6f,0x308", "0x75,0x308,0x2d", "0x61,0x301,0x300", "0x3099", "0x304b,0x3099", "0x30ab,0x309a")
    foreach ($probe in $probes) {
        $s = ""
        foreach ($cp in $probe.Split(",")) { $s += [string][char][int]$cp }
        $rs.AddNew()
        $rs.Fields.Item("Ch").Value = $s
        $rs.Update()
    }
    $rs.Close()
}

# A definition longer than one page: 150 text columns is about 6 KB.
function Build-Wide($db) {
    $ddl = "CREATE TABLE Wide (Id AUTOINCREMENT PRIMARY KEY"
    for ($k = 1; $k -le 150; $k++) { $ddl += ", Col" + $k.ToString("000") + " TEXT(20)" }
    Invoke-Sql $db ($ddl + ")") $dbFailOnError
    $rs = $db.OpenRecordset("Wide", $dbOpenTable)
    for ($r = 1; $r -le 3; $r++) {
        $rs.AddNew()
        for ($k = 1; $k -le 150; $k++) {
            if (($k + $r) % 4 -ne 0) { $rs.Fields.Item("Col" + $k.ToString("000")).Value = "r${r}c${k}" }
        }
        $rs.Update()
    }
    $rs.Close()
}

function Dump-Table($db, [string]$name) {
    $rs = $db.OpenRecordset("SELECT * FROM [$name] ORDER BY Id", $dbOpenSnapshot)
    $lines = New-Object System.Collections.Generic.List[string]
    while (-not $rs.EOF) {
        $cells = New-Object System.Collections.Generic.List[string]
        for ($f = 0; $f -lt $rs.Fields.Count; $f++) {
            $field = $rs.Fields.Item($f)
            $cells.Add($field.Name + "=" + (Format-Cell $field.Value))
        }
        $lines.Add([string]::Join("`t", $cells))
        $rs.MoveNext()
    }
    $rs.Close()
    return [string]::Join("`n", $lines)
}

if ($Command -eq "build-alltypes") { Create-AllTypes $Path }
$engine = New-Object -ComObject DAO.DBEngine.120
if ($Command -eq "compact") {
    # Compacting reads every structure and rebuilds the file; a database
    # the engine cannot make sense of fails here.
    $target = $Path + ".compact.accdb"
    if (Test-Path $target) { Remove-Item $target }
    $engine.CompactDatabase($Path, $target)
    [Console]::Out.Write("ok")
    exit 0
}
$db = $engine.OpenDatabase($Path)
try {
    switch ($Command) {
        "build-alltypes" {
            Build-AllTypes $db $Rows
            Build-Wide $db
            [Console]::Out.Write("ok")
        }
        "build-simple" {
            # A small table for byte-level comparison of single edits.
            Invoke-Sql $db "CREATE TABLE Simple (Id AUTOINCREMENT PRIMARY KEY, N LONG, T TEXT(50))" $dbFailOnError
            Invoke-Sql $db "CREATE INDEX IX_N ON Simple (N)" $dbFailOnError
            for ($i = 1; $i -le 5; $i++) {
                Invoke-Sql $db "INSERT INTO Simple (N, T) VALUES ($($i * 10), 'row $i')" $dbFailOnError
            }
            [Console]::Out.Write("ok")
        }
        "insert-simple" {
            Invoke-Sql $db "INSERT INTO Simple (N, T) VALUES ($Rows, 'inserted $Rows')" $dbFailOnError
            [Console]::Out.Write("ok")
        }
        "delete-simple" {
            Invoke-Sql $db "DELETE FROM Simple WHERE Id = $Rows" $dbFailOnError
            [Console]::Out.Write("ok")
        }
        "insert-alltypes-more" {
            # Rows the engine adds after pyOpenVBA wrote: proof it can still
            # work with the structures.
            for ($i = 1; $i -le $Rows; $i++) {
                Invoke-Sql $db "INSERT INTO AllTypes (Flag, Tiny, Txt) VALUES (TRUE, $i, 'engine after pyopenvba $i')" $dbFailOnError
            }
            [Console]::Out.Write("ok")
        }
        "build-collation" {
            Build-Collation $db
            [Console]::Out.Write("ok")
        }
        "dump" {
            $text = Dump-Table $db $Table
            [Console]::Out.Write((ConvertTo-Json -InputObject $text -Compress))
        }
        default { throw "unknown command $Command" }
    }
} finally {
    $db.Close()
}
