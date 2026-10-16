# _PdfLib.Write.ps1 - PDF writing: page extraction, merge, split, rotate, text->PDF.
# Pure PowerShell/.NET. Rebuilds well-formed output files from parsed objects
# (never edits bytes in place), so xref offsets are always recomputed.
# Requires _PdfLib.Read.ps1 to be loaded first.

Set-StrictMode -Version 2.0

function ConvertTo-PdfNumber {
    # Invariant-culture short number for PDF output (no locale commas).
    param([double]$Value, [int]$Decimals = 2)
    return [Math]::Round($Value, $Decimals).ToString([System.Globalization.CultureInfo]::InvariantCulture)
}

function Escape-PdfLiteralText {
    # Returns @{ Hex=$true/$false; Value=<string> } - hex strings for non-latin1 text.
    param([string]$Text)
    $needsHex = $false
    foreach ($ch in $Text.ToCharArray()) { if ([int]$ch -gt 255) { $needsHex = $true; break } }
    if ($needsHex) {
        $bytes = [System.Text.Encoding]::BigEndianUnicode.GetBytes($Text)
        $sb = New-Object System.Text.StringBuilder
        foreach ($b in $bytes) { [void]$sb.AppendFormat('{0:X2}', $b) }
        return @{ Hex = $true; Value = $sb.ToString() }
    }
    $t = $Text -replace '\\', '\\\\' -replace '\(', '\\(' -replace '\)', '\\)'
    return @{ Hex = $false; Value = $t }
}

function Get-PdfTextPdfStream {
    # Builds a simple content stream drawing wrapped text lines. Returns byte[].
    param([string[]]$Lines, [double]$PageW, [double]$PageH,
          [double]$Margin, [double]$FontSize, [double]$LineHeightFactor = 1.55)

    $lh = $FontSize * $LineHeightFactor
    $maxW = $PageW - 2 * $Margin
    $nl = [System.Environment]::NewLine
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append("BT$nl")
    $y = $PageH - $Margin - $FontSize
    foreach ($line in $Lines) {
        if ($y -lt $Margin) { break }
        $esc = Escape-PdfLiteralText $line
        $x = $Margin
        $op = '1 0 0 1 ' + (ConvertTo-PdfNumber $x) + ' ' + (ConvertTo-PdfNumber $y) + " Tm$nl"
        $op += '/F1 ' + (ConvertTo-PdfNumber $FontSize 1) + " Tf$nl"
        if ($esc.Hex) { $op += '<' + $esc.Value + ">$nl" }
        else { $op += '(' + $esc.Value + ")$nl" }
        $op += "Tj$nl"
        [void]$sb.Append($op)
        $y -= $lh
    }
    [void]$sb.Append("ET$nl")
    return [System.Text.Encoding]::GetEncoding(28591).GetBytes($sb.ToString())
}

function ConvertFrom-PdfWrapText {
    # Word-wraps plain text (LF separated) into display lines for the text PDF.
    param([string]$Text, [double]$MaxWidth, [double]$FontSize)
    $out = New-Object System.Collections.ArrayList
    $rawLines = ($Text -replace "`r`n", "`n" -replace "`r", "`n") -split "`n"
    foreach ($raw in $rawLines) {
        $line = $raw -replace "`t", '    '
        if ([string]::IsNullOrEmpty($line)) { [void]$out.Add(''); continue }
        $cur = ''
        foreach ($word in ($line -split ' ')) {
            $cand = if ($cur) { $cur + ' ' + $word } else { $word }
            while ((Get-PdfTextWidth $cand $FontSize) -gt $MaxWidth -and $word.Length -gt 0) {
                # word itself too long: hard-break it
                if (-not $cur) {
                    $take = ''
                    foreach ($ch in $word.ToCharArray()) {
                        $cand2 = $take + $ch
                        if ((Get-PdfTextWidth $cand2 $FontSize) -gt $MaxWidth) { break }
                        $take = $cand2
                    }
                    if ($take) { [void]$out.Add($take); $word = $word.Substring($take.Length) }
                    else { [void]$out.Add($word.Substring(0,1)); $word = $word.Substring(1) }
                    $cand = $word
                    if (-not $word) { $cand = ''; break }
                } else {
                    [void]$out.Add($cur); $cur = ''; $cand = $word
                }
            }
            $cur = $cand
        }
        if ($null -ne $cur) { [void]$out.Add($cur) }
    }
    return $out.ToArray()
}

function Write-PdfFile {
    # Writes a PDF from prepared objects. $Objects: array of
    # @{ Num=int; Header=string; RawStream=byte[] (optional) }; $RootNum = catalog num.
    param([string]$OutputPath, [object[]]$Objects, [int]$RootNum, [int]$InfoNum = 0)

    $enc = [System.Text.Encoding]::GetEncoding(28591)
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append("%PDF-1.4`n%\xE2\xE3\xCF\xD3`n")
    $offsets = @{}
    $base = $enc.GetBytes($sb.ToString()).Length
    # latin1 is lossless 1 byte = 1 char, so StringBuilder.Length == byte length
    $pos = $base
    $body = New-Object System.Text.StringBuilder
    foreach ($o in ($Objects | Sort-Object { $_.Num })) {
        $offsets[$o.Num] = $pos
        if ($o.RawStream) {
            $head = "$($o.Num) 0 obj`n$($o.Header)`nstream`n"
            [void]$body.Append($head)
            $pos += $head.Length
            $streamText = $enc.GetString([byte[]]$o.RawStream)
            [void]$body.Append($streamText)
            $pos += $streamText.Length
            [void]$body.Append("`nendstream`nendobj`n")
            $pos += 18 # "`nendstream`nendobj`n".Length
            continue
        }
        $objText = "$($o.Num) 0 obj`n$($o.Header)`nendobj`n"
        [void]$body.Append($objText)
        $pos += $objText.Length
    }
    $bodyBytes = $enc.GetBytes($body.ToString())

    $maxNum = 0
    foreach ($o in $Objects) { if ($o.Num -gt $maxNum) { $maxNum = $o.Num } }
    $count = $maxNum + 1
    $xr = New-Object System.Text.StringBuilder
    [void]$xr.AppendLine('xref')
    [void]$xr.AppendLine("0 $count")
    [void]$xr.AppendLine('0000000000 65535 f ')
    for ($i = 1; $i -lt $count; $i++) {
        $off = 0
        if ($offsets.ContainsKey($i)) { $off = [int]$offsets[$i] }
        [void]$xr.Append('{0:D10} 00000 n ' -f $off)
        if ($i -lt $count) { [void]$xr.Append("`n") }
    }
    $trailer = "trailer`n<< /Size $count /Root $RootNum 0 R"
    if ($InfoNum -gt 0) { $trailer += " /Info $InfoNum 0 R" }
    $trailer += " >>`nstartxref`n$($base + $bodyBytes.Length)`n%%EOF`n"

    $fs = [System.IO.File]::Create($OutputPath)
    try {
        $headBytes = $enc.GetBytes($sb.ToString())
        $fs.Write($headBytes, 0, $headBytes.Length)
        $fs.Write($bodyBytes, 0, $bodyBytes.Length)
        $xrBytes = $enc.GetBytes($xr.ToString())
        $fs.Write($xrBytes, 0, $xrBytes.Length)
        $trBytes = $enc.GetBytes($trailer)
        $fs.Write($trBytes, 0, $trBytes.Length)
    } finally { $fs.Dispose() }
}

function New-PdfDocBuilder {
    # Manages lazy copying of objects from source PDFs into a new document.
    # Returns a builder hashtable; use Add-PdfPage / Complete-PdfDoc.
    param()
    return @{
        NextNum   = 3            # 1=catalog 2=pages
        Queue     = New-Object System.Collections.ArrayList   # pending old refs
        Map       = @{}          # "srcIdx:oldNum" -> newNum
        Objects   = New-Object System.Collections.ArrayList   # finished @{Num;Header;RawStream}
        PageNums  = New-Object System.Collections.ArrayList
        Sources   = @()          # array of @{ Objects=hashtable; Idx=int }
    }
}

function Add-PdfSource {
    param($Builder, [hashtable]$Objects)
    $idx = $Builder.Sources.Count
    $Builder.Sources += ,@{ Objects = $Objects; Idx = $idx }
    return $idx
}

function Get-PdfCopiedNum {
    # Lazily allocates a new object number for a source object and queues its copy.
    param($Builder, [int]$SrcIdx, [int]$OldNum)
    $key = "$SrcIdx`:$OldNum"
    if ($Builder.Map.ContainsKey($key)) { return $Builder.Map[$key] }
    $src = $Builder.Sources[$SrcIdx]
    if (-not $src.Objects.ContainsKey($OldNum)) {
        throw "PDF object $OldNum not found in source $($SrcIdx + 1) (file may be encrypted or damaged)"
    }
    $newNum = $Builder.NextNum
    $Builder.NextNum++
    $Builder.Map[$key] = $newNum
    [void]$Builder.Queue.Add(@{ SrcIdx = $SrcIdx; OldNum = $OldNum; NewNum = $newNum })
    return $newNum
}

function Repair-PdfHeader {
    # Cleans a copied object header: drop /Parent, fix /Length for streams.
    param([string]$Header, [bool]$IsStream, [int]$StreamLen)
    $h = $Header
    $h = [regex]::Replace($h, '/Parent\s+\d+\s+\d+\s+R', '')
    if ($IsStream) {
        if ($h -match '/Length\s+\d+\s+\d+\s+R') {
            $h = [regex]::Replace($h, '/Length\s+\d+\s+\d+\s+R', "/Length $StreamLen")
        } elseif ($h -match '/Length\s+\d+') {
            $h = [regex]::Replace($h, '/Length\s+\d+', "/Length $StreamLen")
        } else {
            $h = $h -replace '>>\s*$', " /Length $StreamLen >>"
        }
    }
    return $h
}

function Update-PdfRefsInHeader {
    # Remaps every "N G R" reference in a header through the lazy copier.
    # /Parent is left untouched (stripped elsewhere) so source /Pages nodes are
    # never copied into the new file.
    param($Builder, [int]$SrcIdx, [string]$Header)
    $rx = [regex]'(?<key>/[A-Za-z0-9_]+)\s*(?<ref>\d+)\s+\d+\s+R'
    return $rx.Replace($Header, {
        param($m)
        if ($m.Groups['key'].Value -eq '/Parent') { return $m.Value }
        $newNum = Get-PdfCopiedNum $Builder $SrcIdx ([int]$m.Groups['ref'].Value)
        return "$($m.Groups['key'].Value) $newNum 0 R"
    })
}

function Add-PdfPage {
    # Copies one parsed page (with its content and referenced resources) into the doc.
    param($Builder, [int]$SrcIdx, $Page, [int]$RotateDelta = 0, [double[]]$ForceMediaBox = $null)

    $src = $Builder.Sources[$SrcIdx]
    $header = $Page.RawHeader

    # resolve effective MediaBox (parsed value already includes inheritance)
    $mb = $Page.MediaBox
    if ($ForceMediaBox -and $ForceMediaBox.Count -eq 4) { $mb = $ForceMediaBox }
    if (-not $mb) { $mb = @(0.0, 0.0, 612.0, 792.0) }
    $mbStr = '/MediaBox [ ' + (ConvertTo-PdfNumber $mb[0]) + ' ' + (ConvertTo-PdfNumber $mb[1]) + ' ' +
             (ConvertTo-PdfNumber $mb[2]) + ' ' + (ConvertTo-PdfNumber $mb[3]) + ' ]'
    if ($header -match '/MediaBox') {
        $header = [regex]::Replace($header, '/MediaBox\s*\[[^\]]*\]', $mbStr)
    } else {
        $header = $header -replace '>>\s*$', " $mbStr >>"
    }

    # rotation
    if ($RotateDelta -ne 0) {
        $old = 0
        if ($header -match '/Rotate\s+(\d+)') { $old = [int]$Matches[1] }
        $new = (($old + $RotateDelta) % 360 + 360) % 360
        if ($header -match '/Rotate\s+\d+') {
            $header = [regex]::Replace($header, '/Rotate\s+\d+', "/Rotate $new")
        } else {
            $header = $header -replace '>>\s*$', " /Rotate $new >>"
        }
    }

    # inherited resources: make them explicit so the page is self-contained
    if ($header -notmatch '/Resources' -and ($null -ne $Page.ResourcesRef -or $null -ne $Page.ResourcesText)) {
        if ($null -ne $Page.ResourcesRef) {
            $header = $header -replace '>>\s*$', " /Resources $(Get-PdfCopiedNum $Builder $SrcIdx ([int]$Page.ResourcesRef)) 0 R >>"
        } else {
            $header = $header -replace '>>\s*$', " /Resources $($Page.ResourcesText) >>"
        }
    }

    # drop the parent link first so the source /Pages tree is never copied,
    # then remap every remaining reference (contents, resources, annots, ...)
    $header = [regex]::Replace($header, '/Parent\s+\d+\s+\d+\s+R', '')
    $header = Update-PdfRefsInHeader $Builder $SrcIdx $header

    $newNum = $Builder.NextNum; $Builder.NextNum++
    [void]$Builder.Objects.Add(@{ Num = $newNum; Header = $header.Trim(); RawStream = $null })
    [void]$Builder.PageNums.Add($newNum)
    return $newNum
}

function Complete-PdfDoc {
    # Drains the copy queue, then writes catalog/pages/output. Returns output path.
    param($Builder, [string]$OutputPath)

    $qi = 0
    while ($qi -lt $Builder.Queue.Count) {
        $job = $Builder.Queue[$qi]; $qi++
        $src = $Builder.Sources[$job.SrcIdx]
        $o = $src.Objects[[int]$job.OldNum]
        $h = $o.Header
        $h = Update-PdfRefsInHeader $Builder $job.SrcIdx $h
        $h = [regex]::Replace($h, '/Parent\s+\d+\s+\d+\s+R', '')
        if ($o.IsStream) {
            $h = Repair-PdfHeader $h $true $o.RawStream.Length
            [void]$Builder.Objects.Add(@{ Num = [int]$job.NewNum; Header = $h.Trim(); RawStream = [byte[]]$o.RawStream })
        } else {
            [void]$Builder.Objects.Add(@{ Num = [int]$job.NewNum; Header = $h.Trim(); RawStream = $null })
        }
    }

    $pagesNum = 2
    $kids = ($Builder.PageNums | ForEach-Object { "$_ 0 R" }) -join ' '
    $pagesHeader = "<< /Type /Pages /Count $($Builder.PageNums.Count) /Kids [ $kids ] >>"
    [void]$Builder.Objects.Add(@{ Num = $pagesNum; Header = $pagesHeader; RawStream = $null })
    foreach ($pn in $Builder.PageNums) {
        $obj = $null
        foreach ($o in $Builder.Objects) { if ($o.Num -eq [int]$pn) { $obj = $o; break } }
        if ($obj) {
            if ($obj.Header -match '/Parent') {
                $obj.Header = [regex]::Replace($obj.Header, '/Parent\s+\d+\s+\d+\s+R', "/Parent $pagesNum 0 R")
            } else {
                $obj.Header = $obj.Header -replace '>>\s*$', " /Parent $pagesNum 0 R >>"
            }
        }
    }
    $catalog = @{ Num = 1; Header = '<< /Type /Catalog /Pages 2 0 R >>'; RawStream = $null }
    $all = @($catalog) + @($Builder.Objects.ToArray())
    Write-PdfFile -OutputPath $OutputPath -Objects $all -RootNum 1
    return $OutputPath
}

function ConvertTo-PdfExtract {
    param([string[]]$Paths, [object]$Pages, [string]$OutputPath)
    $builder = New-PdfDocBuilder
    $bytes = [System.IO.File]::ReadAllBytes($Paths[0])
    if (-not (Test-PdfNotEncrypted $bytes)) { throw "PDF is encrypted: '$($Paths[0])'. Offline toolkit cannot decrypt it." }
    $objects = Get-PdfRawObjects $bytes
    $idx = Add-PdfSource $builder $objects
    $pageList = @(Get-PdfPageList $objects)
    $total = @($pageList).Count
    $sel = @(Resolve-PdfPageSelection $Pages $total)
    foreach ($p in $sel) { [void](Add-PdfPage $builder $idx @($pageList[$p - 1])) }
    return (Complete-PdfDoc $builder $OutputPath), @($sel).Count
}

function ConvertTo-PdfMerge {
    param([string[]]$Paths, [string]$OutputPath)
    $builder = New-PdfDocBuilder
    $count = 0
    foreach ($p in $Paths) {
        $bytes = [System.IO.File]::ReadAllBytes($p)
        if (-not (Test-PdfNotEncrypted $bytes)) { throw "PDF is encrypted: '$p'. Offline toolkit cannot decrypt it." }
        $objects = Get-PdfRawObjects $bytes
        $idx = Add-PdfSource $builder $objects
        foreach ($page in @(Get-PdfPageList $objects)) {
            [void](Add-PdfPage $builder $idx $page)
            $count++
        }
    }
    return (Complete-PdfDoc $builder $OutputPath), $count
}

function ConvertTo-PdfSplit {
    param([string]$Path, [int]$PerChunk, [string]$OutputDir, [string]$BaseName)
    $bytes = [System.IO.File]::ReadAllBytes($Path)
    if (-not (Test-PdfNotEncrypted $bytes)) { throw "PDF is encrypted: '$Path'. Offline toolkit cannot decrypt it." }
    $objects = Get-PdfRawObjects $bytes
    $pageList = @(Get-PdfPageList $objects)
    $total = @($pageList).Count
    $dir = New-OutputDir $OutputDir
    $made = @()
    $chunk = 0
    for ($start = 1; $start -le $total; $start += $PerChunk) {
        $chunk++
        $builder = New-PdfDocBuilder
        $idx = Add-PdfSource $builder $objects
        $end = [Math]::Min($start + $PerChunk - 1, $total)
        for ($p = $start; $p -le $end; $p++) { [void](Add-PdfPage $builder $idx @($pageList[$p - 1])) }
        $name = '{0}-part{1:D3}.pdf' -f $BaseName, $chunk
        $out = Join-Path $dir $name
        [void](Complete-PdfDoc $builder $out)
        $made += $out
    }
    return @($made, $total)
}

function ConvertTo-PdfRotate {
    param([string]$Path, [object]$Pages, [int]$Delta, [string]$OutputPath)
    $bytes = [System.IO.File]::ReadAllBytes($Path)
    if (-not (Test-PdfNotEncrypted $bytes)) { throw "PDF is encrypted: '$Path'. Offline toolkit cannot decrypt it." }
    $objects = Get-PdfRawObjects $bytes
    $builder = New-PdfDocBuilder
    $idx = Add-PdfSource $builder $objects
    $pageList = @(Get-PdfPageList $objects)
    $total = @($pageList).Count
    $sel = @(Resolve-PdfPageSelection $Pages $total)
    $selSet = @{}
    foreach ($p in $sel) { $selSet[$p] = $true }
    for ($p = 1; $p -le $total; $p++) {
        $d = 0
        if ($selSet.ContainsKey($p)) { $d = $Delta }
        [void](Add-PdfPage $builder $idx @($pageList[$p - 1]) -RotateDelta $d)
    }
    return (Complete-PdfDoc $builder $OutputPath), $total
}

function ConvertTo-PdfFromText {
    # Creates a simple text PDF. CJK text uses the non-embedded STSong-Light
    # (UniGB-UCS2-H) standard CID font - viewers with Asian font support show it.
    param([string]$Text, [string]$OutputPath,
          [string]$PageSize = 'A4', [double]$FontSize = 11, [double]$Margin = 56.7)

    $size = Get-PdfPageSize $PageSize
    $pageW = $size[0]; $pageH = $size[1]
    $lines = @(ConvertFrom-PdfWrapText $Text ($pageW - 2 * $Margin) $FontSize)

    $perPage = [Math]::Max(1, [Math]::Floor(($pageH - 2 * $Margin) / ($FontSize * 1.55)))
    $builder = New-PdfDocBuilder
    $fontNum = 3; $builder.NextNum = 4
    $hasCjk = $false
    foreach ($ch in $Text.ToCharArray()) { if ([int]$ch -gt 255) { $hasCjk = $true; break } }
    if ($hasCjk) {
        $fontHeader = '<< /Type /Font /Subtype /Type0 /BaseFont /STSong-Light /Encoding /UniGB-UCS2-H ' +
            '/DescendantFonts [ << /Type /Font /Subtype /CIDFontType0 /BaseFont /STSong-Light ' +
            '/CIDSystemInfo << /Registry (Adobe) /Ordering (GB1) /Supplement 4 >> /DW 1000 >> ] >>'
    } else {
        $fontHeader = '<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>'
    }
    [void]$builder.Objects.Add(@{ Num = $fontNum; Header = $fontHeader; RawStream = $null })

    $pageCount = 0
    for ($start = 0; $start -lt $lines.Count; $start += $perPage) {
        $pageCount++
        $slice = @()
        for ($i = $start; $i -lt [Math]::Min($start + $perPage, $lines.Count); $i++) { $slice += $lines[$i] }
        $content = Get-PdfTextPdfStream -Lines $slice -PageW $pageW -PageH $pageH -Margin $Margin -FontSize $FontSize
        $contentNum = $builder.NextNum; $builder.NextNum++
        [void]$builder.Objects.Add(@{ Num = $contentNum; Header = "<< /Length $($content.Length) >>"; RawStream = $content })
        $pageNum = $builder.NextNum; $builder.NextNum++
        $resHdr = "<< /Type /Page /MediaBox [ 0 0 $(ConvertTo-PdfNumber $pageW) $(ConvertTo-PdfNumber $pageH) ] " +
            "/Resources << /Font << /F1 $fontNum 0 R >> >> /Contents $contentNum 0 R >>"
        [void]$builder.Objects.Add(@{ Num = $pageNum; Header = $resHdr; RawStream = $null })
        [void]$builder.PageNums.Add($pageNum)
    }
    if ($pageCount -eq 0) {
        # empty text -> one blank page
        $contentNum = $builder.NextNum; $builder.NextNum++
        [void]$builder.Objects.Add(@{ Num = $contentNum; Header = '<< /Length 8 >>'; RawStream = [System.Text.Encoding]::GetEncoding(28591).GetBytes('BT ET' + "`n") })
        $pageNum = $builder.NextNum; $builder.NextNum++
        [void]$builder.Objects.Add(@{ Num = $pageNum; Header = "<< /Type /Page /MediaBox [ 0 0 $(ConvertTo-PdfNumber $pageW) $(ConvertTo-PdfNumber $pageH) ] /Contents $contentNum 0 R >>"; RawStream = $null })
        [void]$builder.PageNums.Add($pageNum)
    }
    return (Complete-PdfDoc $builder $OutputPath), $pageCount
}

function Get-PdfPageSize {
    param([string]$Name)
    switch ($Name.ToUpperInvariant()) {
        'A4'     { return @([double]595.28, [double]841.89) }
        'A3'     { return @([double]841.89, [double]1190.55) }
        'A5'     { return @([double]419.53, [double]595.28) }
        'LETTER' { return @([double]612.0, [double]792.0) }
        'LEGAL'  { return @([double]612.0, [double]1008.0) }
        default  { throw "Unknown page size '$Name'. Use A4/A3/A5/Letter/Legal." }
    }
}

function Resolve-PdfPageSelection {
    # "1,3,5-8" (or int[]) -> sorted distinct page numbers within 1..total.
    # Empty/null/"all" selects every page. Returns int[].
    param([object]$Pages, [int]$Total)
    $set = New-Object 'System.Collections.Generic.HashSet[int]'
    $hasInput = $false
    if ($null -ne $Pages -and @($Pages).Count -gt 0) {
        $s0 = [string]$Pages[0]
        if (-not [string]::IsNullOrWhiteSpace($s0) -and $s0.Trim().ToLowerInvariant() -ne 'all') {
            $hasInput = $true
            foreach ($item in @($Pages)) {
                $spec = [string]$item
                if ($spec -match '^\s*(\d+)\s*-\s*(\d+)\s*$') {
                    for ($i = [int]$Matches[1]; $i -le [int]$Matches[2]; $i++) { [void]$set.Add($i) }
                } elseif ($spec -match '^\s*\d+\s*$') {
                    [void]$set.Add([int]$spec)
                } elseif ($spec -match '^\s*[\d,\-\s]+\s*$') {
                    foreach ($part in ($spec -split ',')) {
                        $part = $part.Trim()
                        if ($part -match '^(\d+)\s*-\s*(\d+)$') {
                            for ($i = [int]$Matches[1]; $i -le [int]$Matches[2]; $i++) { [void]$set.Add($i) }
                        } elseif ($part -match '^\d+$') { [void]$set.Add([int]$part) }
                        elseif ($part) { throw "Invalid page spec: '$part' (use e.g. 1,3,5-8)" }
                    }
                } else {
                    throw "Invalid page spec: '$spec' (use e.g. 1,3,5-8)"
                }
            }
        }
    }
    if (-not $hasInput) { foreach ($i in 1..$Total) { [void]$set.Add($i) } }
    if ($Total -le 0) { throw 'Document has no pages' }
    $out = @($set | Where-Object { $_ -ge 1 -and $_ -le $Total } | Sort-Object)
    if ($out.Count -eq 0) { throw "No valid pages in selection (document has $Total pages)" }
    return $out
}

function Test-PdfNotEncrypted {
    param([byte[]]$PdfBytes)
    $head = [System.Text.Encoding]::GetEncoding(28591).GetString(
        $PdfBytes, 0, [Math]::Min($PdfBytes.Length, 4096))
    if ($head -match '/Encrypt\s+\d+\s+\d+\s+R') { return $false }
    $tailStart = [Math]::Max(0, $PdfBytes.Length - 8192)
    $tail = [System.Text.Encoding]::GetEncoding(28591).GetString(
        $PdfBytes, $tailStart, $PdfBytes.Length - $tailStart)
    if ($tail -match '/Encrypt\s+\d+\s+\d+\s+R') { return $false }
    return $true
}

function Get-PdfPageCountFast {
    param([string]$Path)
    $bytes = [System.IO.File]::ReadAllBytes($Path)
    if (-not (Test-PdfNotEncrypted $bytes)) { return -1 }
    $objects = Get-PdfRawObjects $bytes
    return (Get-PdfPageList $objects).Count
}
