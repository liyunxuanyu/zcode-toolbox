# Edit-Xlsx.ps1 - Edit Excel .xlsx without Excel installed (OOXML via .NET).
# Usage: powershell -NoProfile -ExecutionPolicy Bypass -File Edit-Xlsx.ps1 -Action <action> ...
# Actions: listnames | getcells | setcells | export | import | addrows | newsheet | delsheet

param(
    [Parameter(Mandatory = $true)][string]$Action,
    [string]$Path = '',
    [int]$Sheet = 1,
    [string]$SheetName = '',
    [string]$InputPath = '',
    [string]$OutputPath = '',
    [string[]]$Cells = @(),        # "A1=value" items
    [string[]]$CellTypes = @(),    # "A1=number|text|bool|formula" (default: auto)
    [string]$Encoding = 'utf8',
    [switch]$NoHeader,
    [switch]$Replace
)

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot '_Common.ps1')

$NS_S = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
$NS_R_PR = 'http://schemas.openxmlformats.org/package/2006/relationships'
$NS_R_OFF = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
$NS_CT = 'http://schemas.openxmlformats.org/package/2006/content-types'

function Resolve-SheetPartPath {
    # Resolves the zip part path of a worksheet by 1-based index or by name.
    param([string]$ZipPath, [int]$SheetNumber, [string]$Name)
    $wbXml = Read-ZipEntryText -ZipPath $ZipPath -EntryName 'xl/workbook.xml'
    if (-not $wbXml) { throw 'Not a valid .xlsx (missing xl/workbook.xml)' }
    $wb = New-XmlDoc $wbXml
    $nsm = Get-XmlNsManager $wb
    $sheets = $wb.SelectNodes('//s:sheet', $nsm)
    if ($sheets.Count -eq 0) { throw 'Workbook has no sheets' }
    $node = $null
    if ($Name) {
        foreach ($sh in $sheets) {
            if ($sh.GetAttribute('name') -eq $Name) { $node = $sh; break }
        }
        if (-not $node) { throw "Sheet '$Name' not found" }
    } else {
        if ($SheetNumber -lt 1 -or $SheetNumber -gt $sheets.Count) {
            throw "Sheet number $SheetNumber out of range 1..$($sheets.Count)"
        }
        $node = $sheets[$SheetNumber - 1]
    }
    $rid = $node.GetAttribute('id', $NS_R_OFF)
    $target = $null
    $relsXml = Read-ZipEntryText -ZipPath $ZipPath -EntryName 'xl/_rels/workbook.xml.rels'
    if ($relsXml -and $rid) {
        $rels = New-XmlDoc $relsXml
        $rnsm = New-Object System.Xml.XmlNamespaceManager($rels.NameTable)
        $rnsm.AddNamespace('pr', $NS_R_PR)
        $rel = $rels.SelectSingleNode("//pr:Relationship[@Id='$rid']", $rnsm)
        if ($rel) {
            $t = $rel.GetAttribute('Target')
            if ($t -notmatch '^/') { $target = 'xl/' + $t } else { $target = $t.TrimStart('/') }
        }
    }
    if (-not $target) { $target = "xl/worksheets/sheet$SheetNumber.xml" }
    if (-not (Test-ZipEntryExists -ZipPath $ZipPath -EntryName $target)) {
        throw "Worksheet part missing in package: $target"
    }
    return $target
}

function Get-SharedStrings {
    # Returns string[] of the shared strings table (empty array when absent).
    param([string]$ZipPath)
    $xml = Read-ZipEntryText -ZipPath $ZipPath -EntryName 'xl/sharedStrings.xml'
    if (-not $xml) { return @() }
    $doc = New-XmlDoc $xml
    $nsm = New-Object System.Xml.XmlNamespaceManager($doc.NameTable)
    $nsm.AddNamespace('s', $NS_S)
    $list = @()
    foreach ($si in $doc.SelectNodes('//s:si', $nsm)) {
        # concatenate all text nodes (rich text runs included)
        $sb = New-Object System.Text.StringBuilder
        foreach ($t in $si.SelectNodes('.//s:t', $nsm)) { [void]$sb.Append($t.InnerText) }
        $list += $sb.ToString()
    }
    return $list
}

function Get-CellValue {
    # Returns the display value of one <c> element.
    param($CellNode, [System.Xml.XmlNamespaceManager]$Nsm, [string[]]$Shared)
    $t = $CellNode.GetAttribute('t')
    if ($t -eq 'inlineStr') {
        $sb = New-Object System.Text.StringBuilder
        foreach ($tn in $CellNode.SelectNodes('.//s:t', $Nsm)) { [void]$sb.Append($tn.InnerText) }
        return $sb.ToString()
    }
    if ($t -eq 's') {
        $v = $CellNode.SelectSingleNode('./s:v', $Nsm)
        if ($v -and $Shared) {
            $idx = [int]$v.InnerText
            if ($idx -ge 0 -and $idx -lt $Shared.Count) { return $Shared[$idx] }
        }
        return ''
    }
    if ($t -eq 'b') {
        $v = $CellNode.SelectSingleNode('./s:v', $Nsm)
        if ($v -and $v.InnerText -eq '1') { return 'TRUE' }
        return 'FALSE'
    }
    $v = $CellNode.SelectSingleNode('./s:v', $Nsm)
    if ($v) { return $v.InnerText }
    return ''
}

# ---------------------------------------------------------------------------
# Sheet row model: parse <sheetData> into row objects, then re-serialize.
# ---------------------------------------------------------------------------

function Get-SheetModel {
    # Splits sheet XML into header / rows / footer and parses rows into cells.
    param([string]$SheetXml)
    $doc = New-XmlDoc $SheetXml
    $nsm = New-Object System.Xml.XmlNamespaceManager($doc.NameTable)
    $nsm.AddNamespace('s', $NS_S)
    $rows = New-Object System.Collections.ArrayList
    $sheetData = $doc.SelectSingleNode('//s:sheetData', $nsm)
    foreach ($rowEl in $sheetData.SelectNodes('./s:row', $nsm)) {
        $cells = New-Object System.Collections.ArrayList
        foreach ($c in $rowEl.SelectNodes('./s:c', $nsm)) {
            $ref = $c.GetAttribute('r')
            $a1 = ConvertFrom-A1 $ref
            [void]$cells.Add(@{
                Ref = $ref; Col = $a1.Col; Row = $a1.Row
                T = $c.GetAttribute('t'); S = $c.GetAttribute('s')
                Inner = $c.InnerXml
            })
        }
        $rowAttrs = @{}
        foreach ($at in $rowEl.Attributes) { if ($at.Name -ne 'r') { $rowAttrs[$at.Name] = $at.Value } }
        [void]$rows.Add(@{ Num = [int]$rowEl.GetAttribute('r'); Attrs = $rowAttrs; Cells = $cells })
    }
    # other interesting root-level bits
    $dim = $doc.SelectSingleNode('//s:dimension/@ref', $nsm)
    return @{
        Xml = $SheetXml
        Doc = $doc
        Nsm = $nsm
        Rows = $rows
        Dimension = if ($dim) { $dim.Value } else { $null }
    }
}

function Copy-XmlIntoDoc {
    # Recreates a parsed fragment node inside $Doc with the given namespace.
    # A fragment string parsed under <root xmlns:s=...> loses the default
    # namespace for unprefixed children (they import as xmlns=""); this rebuilds
    # every element in the worksheet namespace instead.
    param($Doc, $Node, [string]$Ns)
    if ($Node.NodeType -eq [System.Xml.XmlNodeType]::Element) {
        $el = $Doc.CreateElement($Node.LocalName, $Ns)
        foreach ($at in $Node.Attributes) {
            if ($at.Name -eq 'xmlns') { continue } # drop fragment-empty-ns declaration
            [void]$el.SetAttribute($at.Name, $at.Value)
        }
        foreach ($ch in $Node.ChildNodes) {
            [void]$el.AppendChild((Copy-XmlIntoDoc $Doc $ch $Ns))
        }
        return $el
    }
    if ($Node.NodeType -in @([System.Xml.XmlNodeType]::Text, [System.Xml.XmlNodeType]::CDATA)) {
        return $Doc.CreateTextNode($Node.Value)
    }
    return $null
}

function ConvertTo-SheetXml {
    # Serializes the model back into sheet XML, updating <dimension>.
    param($Model)
    $doc = $Model.Doc
    $nsm = $Model.Nsm
    $sheetData = $doc.SelectSingleNode('//s:sheetData', $nsm)

    # rebuild rows in document order
    foreach ($old in @($sheetData.SelectNodes('./s:row', $nsm))) { [void]$sheetData.RemoveChild($old) }
    $sorted = $Model.Rows | Sort-Object { $_.Num }
    foreach ($r in $sorted) {
        $rowEl = $doc.CreateElement('row', $NS_S)
        [void]$rowEl.SetAttribute('r', [string]$r.Num)
        foreach ($k in $r.Attrs.Keys) { [void]$rowEl.SetAttribute($k, [string]$r.Attrs[$k]) }
        $cellsSorted = $r.Cells | Sort-Object { $_.Col }
        foreach ($c in $cellsSorted) {
            $cEl = $doc.CreateElement('c', $NS_S)
            [void]$cEl.SetAttribute('r', $c.Ref)
            if ($c.S) { [void]$cEl.SetAttribute('s', $c.S) }
            if ($c.T) { [void]$cEl.SetAttribute('t', $c.T) }
            if ($c.Inner) {
                $tmp = New-XmlDoc ("<root xmlns:s='$NS_S'>" + $c.Inner + '</root>')
                foreach ($child in @($tmp.DocumentElement.ChildNodes)) {
                    $copied = Copy-XmlIntoDoc $doc $child $NS_S
                    if ($copied) { [void]$cEl.AppendChild($copied) }
                }
            }
            [void]$rowEl.AppendChild($cEl)
        }
        [void]$sheetData.AppendChild($rowEl)
    }

    # refresh dimension
    $maxRow = 0; $maxCol = 0
    foreach ($r in $Model.Rows) {
        if ($r.Num -gt $maxRow) { $maxRow = $r.Num }
        foreach ($c in $r.Cells) { if ($c.Col -gt $maxCol) { $maxCol = $c.Col } }
    }
    $newDim = if ($maxRow -gt 0 -and $maxCol -gt 0) { (ConvertTo-A1 1 1) + ':' + (ConvertTo-A1 $maxRow $maxCol) } else { 'A1' }
    $dimEl = $doc.SelectSingleNode('//s:dimension', $nsm)
    if ($dimEl) { [void]$dimEl.SetAttribute('ref', $newDim) }

    $settings = New-Object System.Xml.XmlWriterSettings
    $settings.Encoding = New-Object System.Text.UTF8Encoding($false)
    $settings.Indent = $false
    $settings.OmitXmlDeclaration = $false
    # serialize through a MemoryStream so the declared encoding (utf-8) matches
    # the bytes; a StringWriter would force an utf-16 declaration
    $ms = New-Object System.IO.MemoryStream
    $xw = [System.Xml.XmlWriter]::Create($ms, $settings)
    try { $doc.Save($xw) } finally { $xw.Dispose() }
    return (New-Object System.Text.UTF8Encoding($false)).GetString($ms.ToArray())
}

function Set-SheetCell {
    # Sets one cell in the model. $Type: auto|number|text|bool|formula
    param($Model, [string]$Ref, [string]$Value, [string]$Type)
    $a1 = ConvertFrom-A1 $Ref
    $rowObj = $null
    foreach ($r in $Model.Rows) { if ($r.Num -eq $a1.Row) { $rowObj = $r; break } }
    if (-not $rowObj) {
        $rowObj = @{ Num = $a1.Row; Attrs = @{}; Cells = (New-Object System.Collections.ArrayList) }
        [void]$Model.Rows.Add($rowObj)
    }
    $existing = $null
    foreach ($c in $rowObj.Cells) { if ($c.Col -eq $a1.Col) { $existing = $c; break } }

    $keepS = if ($existing -and $existing.S) { $existing.S } else { '' }
    $t = ''; $inner = ''
    $effType = $Type
    if ($effType -eq '' -or $null -eq $effType) { $effType = 'auto' }
    if ($effType -eq 'auto') {
        $asNum = 0.0
        $trimmed = $Value.Trim()
        if ($trimmed -match '^=') { $effType = 'formula' }
        elseif ($trimmed -match '^(TRUE|FALSE|真|是|否|假)$') { $effType = 'bool' }
        elseif ([double]::TryParse($trimmed, [System.Globalization.NumberStyles]::Any,
                [System.Globalization.CultureInfo]::InvariantCulture, [ref]$asNum)) { $effType = 'number' }
        else { $effType = 'text' }
    }
    switch ($effType) {
        'number' {
            $v = 0.0
            if (-not [double]::TryParse($Value.Trim(), [System.Globalization.NumberStyles]::Any,
                    [System.Globalization.CultureInfo]::InvariantCulture, [ref]$v)) {
                throw "Cell ${Ref}: '$Value' is not a number"
            }
            $inner = '<v>' + (ConvertTo-PdfNumber6 $v) + '</v>'
        }
        'bool' {
            $tv = $Value.Trim().ToLowerInvariant()
            $bit = if ($tv -in @('true','1','yes','y','是','真')) { '1' } else { '0' }
            $t = 'b'; $inner = "<v>$bit</v>"
        }
        'formula' {
            $f = $Value.TrimStart('=')
            $inner = '<f>' + (ConvertTo-XmlEscapedText $f) + '</f>'
        }
        default {
            $t = 'inlineStr'
            $inner = '<is><t xml:space="preserve">' + (ConvertTo-XmlEscapedText $Value) + '</t></is>'
        }
    }
    $newCell = @{ Ref = (ConvertTo-A1 $a1.Row $a1.Col); Col = $a1.Col; Row = $a1.Row
                  T = $t; S = $keepS; Inner = $inner }
    if ($existing) {
        $idx = $rowObj.Cells.IndexOf($existing)
        $rowObj.Cells[$idx] = $newCell
    } else {
        [void]$rowObj.Cells.Add($newCell)
    }
}

function ConvertTo-PdfNumber6 {
    param([double]$Value)
    if ($Value -eq [Math]::Floor($Value) -and [Math]::Abs($Value) -lt 1e15) {
        return [string][int64]$Value
    }
    return $Value.ToString('R', [System.Globalization.CultureInfo]::InvariantCulture)
}

# ---------------------------------------------------------------------------
# CSV parsing (RFC4180-ish, with quote handling)
# ---------------------------------------------------------------------------

function Read-CsvRows {
    param([string]$CsvPath, [string]$Enc = 'utf8')
    $bytes = [System.IO.File]::ReadAllBytes($CsvPath)
    $text = $null
    $utf8Strict = New-Object System.Text.UTF8Encoding($false, $true)
    if ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) {
        $text = (New-Object System.Text.UTF8Encoding($false)).GetString($bytes, 3, $bytes.Length - 3)
    } elseif ($Enc -match '^(utf8|utf-8)$') {
        try { $text = $utf8Strict.GetString($bytes) }
        catch {
            try { $text = [System.Text.Encoding]::GetEncoding('GB18030').GetString($bytes) }
            catch { $text = [System.Text.Encoding]::Default.GetString($bytes) }
        }
    } else {
        try { $text = [System.Text.Encoding]::GetEncoding($Enc).GetString($bytes) }
        catch { $text = [System.Text.Encoding]::Default.GetString($bytes) }
    }
    $text = $text -replace "`r`n", "`n" -replace "`r", "`n"
    $rows = New-Object System.Collections.ArrayList
    $cur = New-Object System.Collections.ArrayList
    $field = New-Object System.Text.StringBuilder
    $inQuotes = $false
    $i = 0; $n = $text.Length
    while ($i -lt $n) {
        $ch = $text[$i]
        if ($inQuotes) {
            if ($ch -eq '"') {
                if ($i + 1 -lt $n -and $text[$i+1] -eq '"') { [void]$field.Append('"'); $i += 2; continue }
                $inQuotes = $false; $i++; continue
            }
            [void]$field.Append($ch); $i++; continue
        }
        if ($ch -eq '"') { $inQuotes = $true; $i++; continue }
        if ($ch -eq ',') { [void]$cur.Add($field.ToString()); [void]$field.Clear(); $i++; continue }
        if ($ch -eq "`n") {
            [void]$cur.Add($field.ToString()); [void]$field.Clear()
            [void]$rows.Add(@($cur.ToArray()))
            [void]$cur.Clear(); $i++; continue
        }
        [void]$field.Append($ch); $i++
    }
    if ($field.Length -gt 0 -or $cur.Count -gt 0) {
        [void]$cur.Add($field.ToString())
        [void]$rows.Add(@($cur.ToArray()))
    }
    return $rows.ToArray()
}

# ---------------------------------------------------------------------------
# Actions
# ---------------------------------------------------------------------------

function Assert-Workbook {
    Assert-InputFile $Path 'Path'
    if ((Get-Item -LiteralPath $Path).Extension -notin @('.xlsx', '.xlsm')) {
        throw "Expected an .xlsx file, got: '$Path' (for .xls use Excel/WPS to convert first)"
    }
}

switch ($Action.ToLowerInvariant()) {

    'listnames' {
        Assert-Workbook
        $wbXml = Read-ZipEntryText -ZipPath $Path -EntryName 'xl/workbook.xml'
        $wb = New-XmlDoc $wbXml
        $nsm = Get-XmlNsManager $wb
        $sheets = $wb.SelectNodes('//s:sheet', $nsm)
        $i = 0
        foreach ($sh in $sheets) {
            $i++
            $state = $sh.GetAttribute('state')
            if (-not $state) { $state = 'visible' }
            Write-Info ("{0}|{1}|{2}" -f $i, $sh.GetAttribute('name'), $state)
        }
    }

    'getcells' {
        Assert-Workbook
        if ($Cells.Count -eq 0) { throw 'getcells needs -Cells like "A1","B2:D5"' }
        $part = Resolve-SheetPartPath -ZipPath $Path -SheetNumber $Sheet -Name $SheetName
        $shared = @(Get-SharedStrings $Path)
        $shXml = Read-ZipEntryText -ZipPath $Path -EntryName $part
        $model = Get-SheetModel $shXml
        # build lookup
        $lookup = @{}
        foreach ($r in $model.Rows) {
            foreach ($c in $r.Cells) { $lookup[$c.Ref] = $c }
        }
        $expand = New-Object System.Collections.ArrayList
        foreach ($spec in (Split-CellRefs $Cells)) {
            if ($spec -match '^([A-Za-z]+\d+):([A-Za-z]+\d+)$') {
                $a = ConvertFrom-A1 $Matches[1]; $b = ConvertFrom-A1 $Matches[2]
                for ($rr = [Math]::Min($a.Row,$b.Row); $rr -le [Math]::Max($a.Row,$b.Row); $rr++) {
                    for ($cc = [Math]::Min($a.Col,$b.Col); $cc -le [Math]::Max($a.Col,$b.Col); $cc++) {
                        [void]$expand.Add((ConvertTo-A1 $rr $cc))
                    }
                }
            } else {
                [void]$expand.Add($spec)
            }
        }
        Write-Info 'cell|value'
        foreach ($ref in $expand) {
            if ($lookup.ContainsKey($ref)) {
                $c = $lookup[$ref]
                # resolve value from model (shared strings need doc-based value)
                $val = ''
                if ($c.T -eq 's' -and $c.Inner -match '<v>(\d+)</v>') {
                    $idx = [int]$Matches[1]
                    if ($idx -lt $shared.Count) { $val = $shared[$idx] }
                } else {
                    $tmp = New-XmlDoc ("<root xmlns:s='$NS_S'><c r=`"$ref`" t=`"$($c.T)`">" + $c.Inner + '</c></root>')
                    $tnsm = New-Object System.Xml.XmlNamespaceManager($tmp.NameTable)
                    $tnsm.AddNamespace('s', $NS_S)
                    $node = $tmp.DocumentElement.FirstChild
                    $val = Get-CellValue $node $tnsm $shared
                }
                $val = $val -replace "`n", '\n'
                Write-Info "$ref|$val"
            } else {
                Write-Info "$ref|"
            }
        }
    }

    'setcells' {
        Assert-Workbook
        if ($Cells.Count -eq 0) { throw 'setcells needs -Cells like "B3=文本","C4=123","D5==SUM(C2:C4)"' }
        $types = @{}
        foreach ($spec in $CellTypes) {
            if ($spec -match '^([^=]+)=(.+)$') { $types[$Matches[1]] = $Matches[2] }
        }
        Backup-IfNeeded -Path $Path | Out-Null
        $part = Resolve-SheetPartPath -ZipPath $Path -SheetNumber $Sheet -Name $SheetName
        $shXml = Read-ZipEntryText -ZipPath $Path -EntryName $part
        $model = Get-SheetModel $shXml
        $specs = Split-CellSpecs $Cells
        foreach ($spec in $specs) {
            if ($spec -notmatch '^([A-Za-z]+\d+)=(.*)$') { throw "Invalid cell spec (use CellRef=Value): '$spec'" }
            $ref = $Matches[1]; $val = $Matches[2]
            $ty = if ($types.ContainsKey($ref)) { $types[$ref] } else { 'auto' }
            Set-SheetCell -Model $model -Ref $ref -Value $val -Type $ty
        }
        [void](Write-ZipEntryText -ZipPath $Path -EntryName $part -Text (ConvertTo-SheetXml $model))
        Write-Info "Updated $(@($specs).Count) cell(s) in $part (backup created)"
    }

    'export' {
        Assert-Workbook
        $part = Resolve-SheetPartPath -ZipPath $Path -SheetNumber $Sheet -Name $SheetName
        $shared = @(Get-SharedStrings $Path)
        $shXml = Read-ZipEntryText -ZipPath $Path -EntryName $part
        $doc = New-XmlDoc $shXml
        $nsm = New-Object System.Xml.XmlNamespaceManager($doc.NameTable)
        $nsm.AddNamespace('s', $NS_S)
        if (-not $OutputPath) { throw 'export needs -OutputPath (a .csv path)' }
        $sw = New-Object System.IO.StreamWriter($OutputPath, $false, (New-Object System.Text.UTF8Encoding($true)))
        try {
            foreach ($row in $doc.SelectNodes('//s:sheetData/s:row', $nsm)) {
                # NB: variable names are case-insensitive; $cellMap avoids the
                # $Cells parameter (a string[]) which IS this same variable
                $cellMap = @{}
                $maxC = 0
                foreach ($c in $row.SelectNodes('./s:c', $nsm)) {
                    $ref = $c.GetAttribute('r')
                    if ($ref -match '^([A-Za-z]+)') {
                        $cn = [string](ConvertTo-ColumnNumber $Matches[1])
                        $cellMap[$cn] = Get-CellValue $c $nsm $shared
                        if ([int]$cn -gt $maxC) { $maxC = [int]$cn }
                    }
                }
                $fields = @()
                for ($cc = 1; $cc -le $maxC; $cc++) {
                    if ($cellMap.ContainsKey([string]$cc)) { $fields += $cellMap[[string]$cc] } else { $fields += '' }
                }
                [void]$sw.WriteLine((ConvertTo-CsvLine $fields))
            }
        } finally { $sw.Dispose() }
        Write-Info "Exported to: $OutputPath"
    }

    'import' {
        Assert-Workbook
        Assert-InputFile $InputPath 'InputPath'
        $rows = @(Read-CsvRows -CsvPath $InputPath -Enc $Encoding)
        if ($NoHeader -eq $false -and @($rows).Count -gt 0) {
            # header row is imported as row 1 like Excel paste
        }
        Backup-IfNeeded -Path $Path | Out-Null
        $part = Resolve-SheetPartPath -ZipPath $Path -SheetNumber $Sheet -Name $SheetName
        $shXml = Read-ZipEntryText -ZipPath $Path -EntryName $part
        $model = Get-SheetModel $shXml
        if (-not $Replace) { $model.Rows.Clear() }
        foreach ($r in $rows) {
            for ($cc = 0; $cc -lt $r.Count; $cc++) {
                $ref = (ConvertTo-A1 1 ($cc + 1))
                Set-SheetCell -Model $model -Ref $ref -Value ([string]$r[$cc]) -Type 'auto'
            }
        }
        [void](Write-ZipEntryText -ZipPath $Path -EntryName $part -Text (ConvertTo-SheetXml $model))
        Write-Info "Imported $($rows.Count) row(s) into $part"
    }

    'addrows' {
        Assert-Workbook
        Assert-InputFile $InputPath 'InputPath'
        $rows = @(Read-CsvRows -CsvPath $InputPath -Enc $Encoding)
        Backup-IfNeeded -Path $Path | Out-Null
        $part = Resolve-SheetPartPath -ZipPath $Path -SheetNumber $Sheet -Name $SheetName
        $shXml = Read-ZipEntryText -ZipPath $Path -EntryName $part
        $model = Get-SheetModel $shXml
        $maxRow = 0
        foreach ($r in $model.Rows) { if ($r.Num -gt $maxRow) { $maxRow = $r.Num } }
        $rn = $maxRow
        foreach ($r in $rows) {
            $rn++
            for ($cc = 0; $cc -lt $r.Count; $cc++) {
                $ref = (ConvertTo-A1 $rn ($cc + 1))
                Set-SheetCell -Model $model -Ref $ref -Value ([string]$r[$cc]) -Type 'auto'
            }
        }
        [void](Write-ZipEntryText -ZipPath $Path -EntryName $part -Text (ConvertTo-SheetXml $model))
        Write-Info "Appended $($rows.Count) row(s) after row $maxRow into $part"
    }

    'newsheet' {
        Assert-Workbook
        if (-not $SheetName) { throw 'newsheet needs -SheetName' }
        if ($SheetName.Length -gt 31) { throw 'Sheet name exceeds 31 characters' }
        if ($SheetName -match '[\[\]\*\?/\\:]') { throw 'Sheet name contains invalid characters: [ ] * ? / \ :' }
        $wbXml = Read-ZipEntryText -ZipPath $Path -EntryName 'xl/workbook.xml'
        $wb = New-XmlDoc $wbXml
        $nsm = Get-XmlNsManager $wb
        $sheets = $wb.SelectNodes('//s:sheet', $nsm)
        foreach ($sh in $sheets) {
            if ($sh.GetAttribute('name') -eq $SheetName) {
                if ($Replace) {
                    & $PSCommandPath -Action delsheet -Path $Path -SheetName $SheetName
                    $wbXml = Read-ZipEntryText -ZipPath $Path -EntryName 'xl/workbook.xml'
                    $wb = New-XmlDoc $wbXml
                    $nsm = Get-XmlNsManager $wb
                    $sheets = $wb.SelectNodes('//s:sheet', $nsm)
                } else {
                    throw "Sheet '$SheetName' already exists (use -Replace to overwrite)"
                }
            }
        }

        # find free part number
        $names = Get-ZipEntryNames -ZipPath $Path
        $partNum = 1
        while (($names -contains "xl/worksheets/sheet$partNum.xml")) { $partNum++ }
        $partName = "xl/worksheets/sheet$partNum.xml"

        # new sheetId and rId
        $maxSheetId = 0; $maxRid = 0
        foreach ($sh in $sheets) {
            $sid = 0
            if ([int]::TryParse($sh.GetAttribute('sheetId'), [ref]$sid)) {
                if ($sid -gt $maxSheetId) { $maxSheetId = $sid }
            }
        }
        $relsXml = Read-ZipEntryText -ZipPath $Path -EntryName 'xl/_rels/workbook.xml.rels'
        if ($relsXml) {
            foreach ($m in [regex]::Matches($relsXml, 'Id="rId(\d+)"')) {
                if ([int]$m.Groups[1].Value -gt $maxRid) { $maxRid = [int]$m.Groups[1].Value }
            }
        }
        $newSid = $maxSheetId + 1; $newRid = "rId$($maxRid + 1)"

        # insert <sheet> into workbook.xml
        $sheetTag = "<sheet name=`"$(ConvertTo-XmlEscapedText $SheetName)`" sheetId=`"$newSid`" r:id=`"$newRid`"/>"
        $newWbXml = $wbXml -replace '</sheets>', "  $sheetTag`n  </sheets>"
        [void](Write-ZipEntryText -ZipPath $Path -EntryName 'xl/workbook.xml' -Text $newWbXml)

        # relationship
        $relTag = "<Relationship Id=`"$newRid`" Type=`"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet`" Target=`"worksheets/sheet$partNum.xml`"/>"
        if ($relsXml) {
            $newRels = $relsXml -replace '</Relationships>', "  $relTag`n  </Relationships>"
        } else {
            $newRels = "<?xml version=`"1.0`" encoding=`"UTF-8`" standalone=`"yes`"?>`n<Relationships xmlns=`"$NS_R_PR`">  $relTag`n</Relationships>"
        }
        [void](Write-ZipEntryText -ZipPath $Path -EntryName 'xl/_rels/workbook.xml.rels' -Text $newRels)

        # content type
        $ctXml = Read-ZipEntryText -ZipPath $Path -EntryName '[Content_Types].xml'
        $ctOverride = "<Override PartName=`"/$partName`" ContentType=`"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml`"/>"
        if ($ctXml -and $ctXml -notmatch [regex]::Escape("/$partName")) {
            $newCt = $ctXml -replace '</Types>', "  $ctOverride`n</Types>"
            [void](Write-ZipEntryText -ZipPath $Path -EntryName '[Content_Types].xml' -Text $newCt)
        }

        # part content: from template empty sheet, then optional import
        $templateSheet = Read-FileTextUtf8 (Join-Path (Split-Path -Parent $PSScriptRoot) 'templates\xlsx-empty-sheet.xml')
        [void](Write-ZipEntryText -ZipPath $Path -EntryName $partName -Text $templateSheet)

        if ($InputPath) {
            Assert-InputFile $InputPath 'InputPath'
            $rows = @(Read-CsvRows -CsvPath $InputPath -Enc $Encoding)
            $shXml = Read-ZipEntryText -ZipPath $Path -EntryName $partName
            $model = Get-SheetModel $shXml
            $model.Rows.Clear()
            $rn = 0
            foreach ($r in $rows) {
                $rn++
                for ($cc = 0; $cc -lt $r.Count; $cc++) {
                    $ref = (ConvertTo-A1 $rn ($cc + 1))
                    Set-SheetCell -Model $model -Ref $ref -Value ([string]$r[$cc]) -Type 'auto'
                }
            }
            [void](Write-ZipEntryText -ZipPath $Path -EntryName $partName -Text (ConvertTo-SheetXml $model))
        }
        Write-Info "Created sheet '$SheetName' ($partName)"
    }

    'delsheet' {
        Assert-Workbook
        if (-not $SheetName) { throw 'delsheet needs -SheetName' }
        $wbXml = Read-ZipEntryText -ZipPath $Path -EntryName 'xl/workbook.xml'
        $wb = New-XmlDoc $wbXml
        $nsm = Get-XmlNsManager $wb
        $sheets = $wb.SelectNodes('//s:sheet', $nsm)
        if ($sheets.Count -le 1) { throw 'Cannot delete the only remaining sheet' }
        $target = $null
        foreach ($sh in $sheets) {
            if ($sh.GetAttribute('name') -eq $SheetName) { $target = $sh; break }
        }
        if (-not $target) { throw "Sheet '$SheetName' not found" }
        $rid = $target.GetAttribute('id', $NS_R_OFF)

        $newWbXml = [regex]::Replace($wbXml,
            [regex]::Escape($target.OuterXml), '')
        # also match self-closing with attribute order variants
        if ($newWbXml -eq $wbXml) {
            $newWbXml = [regex]::Replace($wbXml,
                "<sheet [^>]*name=`"$( [regex]::Escape($SheetName) )`"[^>]*/>", '')
        }
        [void](Write-ZipEntryText -ZipPath $Path -EntryName 'xl/workbook.xml' -Text $newWbXml)

        $relsXml = Read-ZipEntryText -ZipPath $Path -EntryName 'xl/_rels/workbook.xml.rels'
        if ($relsXml -and $rid) {
            $newRels = [regex]::Replace($relsXml, "<Relationship [^>]*Id=`"$rid`"[^>]*/>", '')
            [void](Write-ZipEntryText -ZipPath $Path -EntryName 'xl/_rels/workbook.xml.rels' -Text $newRels)
        }

        # find the part to remove
        $part = $null
        if ($relsXml -and $rid) {
            $m = [regex]::Match($relsXml, "<Relationship [^>]*Id=`"$rid`"[^>]*Target=`"([^`"]+)`"")
            if ($m.Success) {
                $t = $m.Groups[1].Value
                if ($t -notmatch '^/') { $part = 'xl/' + $t } else { $part = $t.TrimStart('/') }
            }
        }
        if ($part -and (Test-ZipEntryExists -ZipPath $Path -EntryName $part)) {
            [void](Remove-ZipEntry -ZipPath $Path -EntryName $part)
            $ctXml = Read-ZipEntryText -ZipPath $Path -EntryName '[Content_Types].xml'
            if ($ctXml) {
                $newCt = [regex]::Replace($ctXml, "<Override PartName=`"/$([regex]::Escape($part))`"[^>]*/>", '')
                if ($newCt -ne $ctXml) {
                    [void](Write-ZipEntryText -ZipPath $Path -EntryName '[Content_Types].xml' -Text $newCt)
                }
            }
        }
        Write-Info "Deleted sheet '$SheetName'"
    }

    default {
        throw "Unknown action '$Action'. Valid: listnames, getcells, setcells, export, import, addrows, newsheet, delsheet"
    }
}
