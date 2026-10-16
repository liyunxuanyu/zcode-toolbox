# _Common.ps1 - shared helpers for offline-office-docs scripts.
# Target: Windows PowerShell 5.1 (.NET Framework 4.x), fully offline.
# All scripts dot-source this file. Encoding: save/read as UTF-8 everywhere.

Set-StrictMode -Version 2.0

# ---------------------------------------------------------------------------
# Paths / logging
# ---------------------------------------------------------------------------

$script:ScriptFile = if ($PSCommandPath) { $PSCommandPath } else { $MyInvocation.MyCommand.Path }
$script:ScriptDir  = Split-Path -Parent $script:ScriptFile

function Write-Info { param([string]$Message) Write-Host $Message }
function Write-Warn2 { param([string]$Message) Write-Host "[WARN] $Message" -ForegroundColor Yellow }
function Write-Err2  { param([string]$Message) Write-Host "[ERROR] $Message" -ForegroundColor Red }

function Test-FileReadable {
    param([string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path)) { return $false }
    return (Test-Path -LiteralPath $Path -PathType Leaf)
}

function Assert-InputFile {
    param([string]$Path, [string]$ParamName = 'Path')
    if (-not (Test-FileReadable $Path)) {
        throw "$ParamName does not exist or is not a file: '$Path'"
    }
}

function New-OutputDir {
    # Creates the output directory (parents included) and returns its full path.
    param([string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path)) { throw "Output path is empty." }
    $full = [System.IO.Path]::GetFullPath($Path)
    if (-not (Test-Path -LiteralPath $full)) {
        New-Item -ItemType Directory -Path $full -Force | Out-Null
    }
    return $full
}

function Backup-IfNeeded {
    # Copies the file to "<name>.bak-<yyyyMMdd-HHmmss>" before first modification.
    param([string]$Path, [int]$Keep = 5)
    $dir = Split-Path -Parent (Resolve-Path -LiteralPath $Path).Path
    $name = Split-Path -Leaf $Path
    $stamp = (Get-Date).ToString('yyyyMMdd-HHmmss')
    $bak = Join-Path $dir ("$name.bak-$stamp")
    Copy-Item -LiteralPath $Path -Destination $bak -Force
    # prune old backups beyond $Keep
    $prefix = "$name.bak-"
    $old = Get-ChildItem -LiteralPath $dir -File -Filter "$prefix*" -ErrorAction SilentlyContinue |
        Sort-Object Name -Descending
    if ($old -and @($old).Count -gt $Keep) {
        $old | Select-Object -Skip $Keep | ForEach-Object { Remove-Item -LiteralPath $_.FullName -Force }
    }
    return $bak
}

# ---------------------------------------------------------------------------
# Text / encoding helpers
# ---------------------------------------------------------------------------

function Read-FileTextUtf8 {
    param([string]$Path)
    $utf8 = New-Object System.Text.UTF8Encoding($false, $true) # throw on invalid
    try {
        return $utf8.GetString([System.IO.File]::ReadAllBytes($Path))
    } catch {
        # fall back to .NET detection (BOM / default ANSI)
        return [System.IO.File]::ReadAllText($Path)
    }
}

function Write-FileTextUtf8 {
    param([string]$Path, [string]$Text)
    $utf8 = New-Object System.Text.UTF8Encoding($false)
    [System.IO.File]::WriteAllText($Path, $Text, $utf8)
}

# ---------------------------------------------------------------------------
# CSV output helpers (agent-friendly, no PS formatting pitfalls)
# ---------------------------------------------------------------------------

function ConvertTo-CsvLine {
    param([string[]]$Fields)
    $sb = New-Object System.Text.StringBuilder
    for ($i = 0; $i -lt $Fields.Count; $i++) {
        if ($i -gt 0) { [void]$sb.Append(',') }
        $f = $Fields[$i]
        if ($null -eq $f) { $f = '' }
        if ($f -match '[",\r\n]') {
            $f = '"' + ($f -replace '"','""') + '"'
        }
        [void]$sb.Append($f)
    }
    return $sb.ToString()
}

function Write-TableReport {
    # Prints a simple aligned text table. $Rows = array of string[] arrays.
    param([string[]]$Header, [object[]]$Rows)
    Write-Info (ConvertTo-CsvLine $Header)
    foreach ($r in $Rows) { Write-Info (ConvertTo-CsvLine @($r)) }
}

# ---------------------------------------------------------------------------
# ZIP helpers (System.IO.Compression) - the core of all OOXML editing
# ---------------------------------------------------------------------------

function Get-ZipEntries {
    param([string]$ZipPath)
    Add-Type -AssemblyName System.IO.Compression
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $zip = [System.IO.Compression.ZipFile]::OpenRead($ZipPath)
    try {
        $list = @()
        foreach ($e in $zip.Entries) {
            $list += ,@($e.FullName, [int64]$e.Length)
        }
        return ,$list
    } finally { $zip.Dispose() }
}

function Read-ZipEntryText {
    # Reads one entry as text. Tries to detect XML and use its declared encoding;
    # falls back to UTF-8. Returns $null when the entry does not exist.
    param([string]$ZipPath, [string]$EntryName)
    Add-Type -AssemblyName System.IO.Compression
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $zip = [System.IO.Compression.ZipFile]::OpenRead($ZipPath)
    try {
        $e = $zip.GetEntry($EntryName)
        if (-not $e) { return $null }
        $s = $e.Open()
        try {
            $ms = New-Object System.IO.MemoryStream
            try {
                $s.CopyTo($ms)
                $bytes = $ms.ToArray()
            } finally { $ms.Dispose() }
        } finally { $s.Dispose() }
    } finally { $zip.Dispose() }

    # BOM check first
    if ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) {
        return (New-Object System.Text.UTF8Encoding($false)).GetString($bytes, 3, $bytes.Length - 3)
    }
    if ($bytes.Length -ge 2 -and (($bytes[0] -eq 0xFF -and $bytes[1] -eq 0xFE) -or ($bytes[0] -eq 0xFE -and $bytes[1] -eq 0xFF))) {
        return [System.Text.Encoding]::Unicode.GetString($bytes)
    }
    # Peek at the XML declaration: <?xml version="1.0" encoding="gb2312" ?>
    $head = [System.Text.Encoding]::ASCII.GetString($bytes, 0, [Math]::Min(200, $bytes.Length))
    if ($head -match 'encoding\s*=\s*"([^"]+)"') {
        $encName = $Matches[1]
        try {
            $enc = [System.Text.Encoding]::GetEncoding($encName)
            return $enc.GetString($bytes)
        } catch { }
    }
    # Default: UTF-8 (OOXML standard)
    return (New-Object System.Text.UTF8Encoding($false)).GetString($bytes)
}

function Read-ZipEntryBytes {
    param([string]$ZipPath, [string]$EntryName)
    Add-Type -AssemblyName System.IO.Compression
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $zip = [System.IO.Compression.ZipFile]::OpenRead($ZipPath)
    try {
        $e = $zip.GetEntry($EntryName)
        if (-not $e) { return $null }
        $s = $e.Open()
        try {
            $ms = New-Object System.IO.MemoryStream
            try {
                $s.CopyTo($ms)
                return $ms.ToArray()
            } finally { $ms.Dispose() }
        } finally { $s.Dispose() }
    } finally { $zip.Dispose() }
}

function Write-ZipEntryText {
    # Creates or replaces one entry with UTF-8 (no BOM) text. Returns $true if the
    # entry existed (replaced), $false if created.
    param([string]$ZipPath, [string]$EntryName, [string]$Text)
    $bytes = (New-Object System.Text.UTF8Encoding($false)).GetBytes($Text)
    return (Write-ZipEntryBytes -ZipPath $ZipPath -EntryName $EntryName -Bytes $bytes)
}

function Write-ZipEntryBytes {
    param([string]$ZipPath, [string]$EntryName, [byte[]]$Bytes)
    Add-Type -AssemblyName System.IO.Compression
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $zip = [System.IO.Compression.ZipFile]::Open($ZipPath, 'Update')
    try {
        $old = $zip.GetEntry($EntryName)
        if ($old) { $old.Delete() }
        $e = $zip.CreateEntry($EntryName, [System.IO.Compression.CompressionLevel]::Optimal)
        $s = $e.Open()
        try { $s.Write($Bytes, 0, $Bytes.Length) } finally { $s.Dispose() }
        return ($null -ne $old)
    } finally { $zip.Dispose() }
}

function Remove-ZipEntry {
    param([string]$ZipPath, [string]$EntryName)
    Add-Type -AssemblyName System.IO.Compression
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $zip = [System.IO.Compression.ZipFile]::Open($ZipPath, 'Update')
    try {
        $e = $zip.GetEntry($EntryName)
        if ($e) { $e.Delete(); return $true }
        return $false
    } finally { $zip.Dispose() }
}

function Get-ZipEntryNames {
    param([string]$ZipPath)
    $entries = Get-ZipEntries -ZipPath $ZipPath
    return @($entries | ForEach-Object { $_[0] })
}

function Test-ZipEntryExists {
    param([string]$ZipPath, [string]$EntryName)
    return ((Get-ZipEntryNames -ZipPath $ZipPath) -contains $EntryName)
}

# ---------------------------------------------------------------------------
# OOXML shared helpers
# ---------------------------------------------------------------------------

function Get-OoxmlPartNames {
    # Convenience: list all parts of an OOXML package that look like document parts.
    param([string]$Path)
    return (Get-ZipEntryNames -ZipPath $Path)
}

function Get-OoxmlTextContent {
    # Extract all <w:t>/<a:t>/<t> text runs from an XML part, preserving order,
    # with paragraph breaks (<w:p>) and cell tabs (<a:br>) as newlines.
    param([string]$Xml)
    $doc = New-Object System.Xml.XmlDocument
    $doc.PreserveWhitespace = $false
    $doc.LoadXml($Xml)
    $nsm = New-Object System.Xml.XmlNamespaceManager($doc.NameTable)
    $nsm.AddNamespace('w',  'http://schemas.openxmlformats.org/wordprocessingml/2006/main')
    $nsm.AddNamespace('a',  'http://schemas.openxmlformats.org/drawingml/2006/main')
    $nsm.AddNamespace('s',  'http://schemas.openxmlformats.org/spreadsheetml/2006/main')
    $nsm.AddNamespace('m',  'http://schemas.openxmlformats.org/officeDocument/2006/math')

    $out = New-Object System.Text.StringBuilder
    $nodes = $doc.SelectNodes('//*[self::w:t or self::a:t or self::s:t or self::m:t]', $nsm)
    $lastPara = $null
    foreach ($n in $nodes) {
        if ($n.ParentNode -and $lastPara -and (-not [object]::ReferenceEquals($n.ParentNode, $lastPara))) {
            [void]$out.Append("`n")
        }
        $lastPara = $n.ParentNode
        [void]$out.Append($n.InnerText)
    }
    return $out.ToString()
}

function New-OoxmlPackageFromTemplate {
    # Copies a template package to the destination (parent dirs are created).
    param([string]$TemplatePath, [string]$DestinationPath)
    $destFull = [System.IO.Path]::GetFullPath($DestinationPath)
    $destDir = Split-Path -Parent $destFull
    if (-not (Test-Path -LiteralPath $destDir)) {
        New-Item -ItemType Directory -Path $destDir -Force | Out-Null
    }
    Copy-Item -LiteralPath $TemplatePath -Destination $destFull -Force
    return $destFull
}

# ---------------------------------------------------------------------------
# XML doc helpers
# ---------------------------------------------------------------------------

function New-XmlDoc {
    param([string]$Xml = $null)
    $doc = New-Object System.Xml.XmlDocument
    $doc.PreserveWhitespace = $true
    if ($Xml) { $doc.LoadXml($Xml) }
    return $doc
}

function Get-XmlNsManager {
    param($Doc)
    $nsm = New-Object System.Xml.XmlNamespaceManager($Doc.NameTable)
    $nsm.AddNamespace('w',  'http://schemas.openxmlformats.org/wordprocessingml/2006/main')
    $nsm.AddNamespace('a',  'http://schemas.openxmlformats.org/drawingml/2006/main')
    $nsm.AddNamespace('s',  'http://schemas.openxmlformats.org/spreadsheetml/2006/main')
    $nsm.AddNamespace('r',  'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
    $nsm.AddNamespace('p',  'http://schemas.openxmlformats.org/presentationml/2006/main')
    $nsm.AddNamespace('cp', 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties')
    $nsm.AddNamespace('dc', 'http://purl.org/dc/elements/1.1/')
    $nsm.AddNamespace('dcterms', 'http://purl.org/dc/terms/')
    $nsm.AddNamespace('ct', 'http://schemas.openxmlformats.org/package/2006/content-types')
    # leading comma prevents PowerShell from unrolling the IEnumerable manager
    # into a string array of its built-in namespaces on return
    return ,$nsm
}

# ---------------------------------------------------------------------------
# Excel: A1 <-> index helpers
# ---------------------------------------------------------------------------

function ConvertTo-ColumnNumber {
    # "A" -> 1, "Z" -> 26, "AA" -> 27
    param([string]$ColumnLetter)
    $n = 0
    foreach ($ch in $ColumnLetter.ToUpperInvariant().ToCharArray()) {
        if ($ch -lt 'A' -or $ch -gt 'Z') { throw "Invalid column letter: '$ColumnLetter'" }
        $n = $n * 26 + ([int]$ch - [int][char]'A' + 1)
    }
    return $n
}

function ConvertTo-ColumnLetters {
    # 1 -> "A", 27 -> "AA"
    param([int]$ColumnNumber)
    if ($ColumnNumber -lt 1) { throw "Column number must be >= 1" }
    $sb = New-Object System.Text.StringBuilder
    $n = $ColumnNumber
    while ($n -gt 0) {
        $rem = ($n - 1) % 26
        [void]$sb.Insert(0, [char]([int][char]'A' + $rem))
        $n = [Math]::Floor(($n - 1) / 26)
    }
    return $sb.ToString()
}

function ConvertFrom-A1 {
    # "B3" -> @{ Row = 3; Col = 2 }
    param([string]$CellRef)
    if ($CellRef -notmatch '^([A-Za-z]+)(\d+)$') { throw "Invalid A1 reference: '$CellRef'" }
    return @{
        Row = [int]$Matches[2]
        Col = ConvertTo-ColumnNumber $Matches[1]
    }
}

function ConvertTo-A1 {
    param([int]$Row, [int]$Col)
    return ((ConvertTo-ColumnLetters $Col) + $Row)
}

# ---------------------------------------------------------------------------
# JSON output (ConvertTo-Json exists in PS 5.1; depth must be generous)
# ---------------------------------------------------------------------------

function ConvertTo-JsonSafe {
    param($Data, [int]$Depth = 10)
    return (ConvertTo-Json -InputObject $Data -Depth $Depth)
}

function ConvertTo-SheetValue {
    # Converts a CSV cell string to a typed xlsx value: number, bool, inline string.
    param([string]$Raw, [string]$Type)
    switch ($Type) {
        'number' {
            $v = 0.0
            $s = $Raw.Trim()
            if ([double]::TryParse($s, [System.Globalization.NumberStyles]::Any,
                    [System.Globalization.CultureInfo]::InvariantCulture, [ref]$v)) {
                if ($v -eq [Math]::Floor($v) -and [Math]::Abs($v) -lt 1e15) {
                    return @{ t = 'n'; v = [string][int64]$v }
                }
                return @{ t = 'n'; v = $s }
            }
            return @{ t = 'inlineStr'; v = $Raw }
        }
        'bool' {
            $t = $Raw.Trim().ToLowerInvariant()
            if ($t -in @('true','1','yes','y','是','真')) { return @{ t = 'b'; v = '1' } }
            if ($t -in @('false','0','no','n','否','假')) { return @{ t = 'b'; v = '0' } }
            return @{ t = 'inlineStr'; v = $Raw }
        }
        default { return @{ t = 'inlineStr'; v = $Raw } }
    }
}

function ConvertTo-XmlEscapedText {
    param([string]$Text)
    return [System.Security.SecurityElement]::Escape($Text)
}

function Split-CellSpecs {
    # Normalizes cell spec lists: when invoked with -File, "A1=x","B1=y" can
    # arrive as ONE string. Splits at commas that precede a CellRef= pattern,
    # so values containing commas survive.
    param([string[]]$Items)
    $joined = (@($Items) | ForEach-Object { ([string]$_).Trim() }) -join ','
    $joined = $joined.Trim('"')
    $parts = [regex]::Split($joined, ',(?=\s*[A-Za-z]{1,3}\d+\s*=)')
    return @($parts | Where-Object { $_ -match '\S' } | ForEach-Object { $_.Trim() })
}

function Split-CellRefs {
    # Same normalization for pure reference lists like "A1,B2:D5".
    param([string[]]$Items)
    $joined = (@($Items) | ForEach-Object { ([string]$_).Trim() }) -join ','
    $joined = $joined.Trim('"')
    $parts = [regex]::Split($joined, ',(?=\s*[A-Za-z]{1,3}\d+)')
    return @($parts | Where-Object { $_ -match '\S' } | ForEach-Object { $_.Trim() })
}
