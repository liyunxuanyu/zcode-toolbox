# _PdfLib.Read.ps1 - PDF parsing: object scan, page list, stream inflate, text extraction.
# Works on PDF 1.x files. Pure PowerShell/.NET, no external tools.
# Load with: . (Join-Path $PSScriptRoot '_PdfLib.Read.ps1')

Set-StrictMode -Version 2.0

function Get-PdfLatin1Text {
    # Lossless binary -> string (1 byte = 1 char) so offsets/indexof stay byte-accurate.
    param([byte[]]$Bytes)
    return [System.Text.Encoding]::GetEncoding(28591).GetString($Bytes)
}

function Expand-PdfFlate {
    param([byte[]]$Bytes)
    if ($null -eq $Bytes -or $Bytes.Length -eq 0) { return @() }
    foreach ($offset in @(0, 2)) {
        if ($Bytes.Length -le $offset) { continue }
        try {
            $ms = New-Object System.IO.MemoryStream(,$Bytes)
            [void]$ms.Seek($offset, 'Begin')
            $ds = New-Object System.IO.Compression.DeflateStream($ms, [System.IO.Compression.CompressionMode]::Decompress)
            $out = New-Object System.IO.MemoryStream
            try { $ds.CopyTo($out) } finally { $ds.Dispose(); $out.Dispose() }
            return $out.ToArray()
        } catch { }
    }
    return $Bytes
}

function Get-PdfRawObjects {
    # Scans every "N G obj ... endobj". Returns hashtable objnum ->
    #   @{ Header=<dict text>; Stream=<byte[] decompressed or $null>; RawStream=<byte[]>;
    #      Filters=<string>; IsStream=<bool> }
    param([byte[]]$PdfBytes)

    $text = Get-PdfLatin1Text $PdfBytes
    $objects = @{}
    $rx = [regex]'\b(\d+)\s+\d+\s+obj\b'
    $pos = 0
    while ($true) {
        $m = $rx.Match($text, $pos)
        if (-not $m.Success) { break }
        $objNum = [int]$m.Groups[1].Value
        $bodyStart = $m.Index + $m.Length
        $endIdx = $text.IndexOf('endobj', $bodyStart)
        if ($endIdx -lt 0) { $endIdx = $text.Length }
        $body = $text.Substring($bodyStart, $endIdx - $bodyStart)

        $header = $body
        $rawStream = $null
        $stream = $null
        $filters = ''
        $sIdx = $body.IndexOf(' stream')
        if ($sIdx -lt 0) { $sIdx = $body.IndexOf('stream') }
        $dictEnd = $body.LastIndexOf('>>')
        if ($sIdx -ge 0 -and $dictEnd -lt $sIdx) {
            # dict ends before 'stream' -> this object carries a stream
            $header = $body.Substring(0, $sIdx)
            $ds = $sIdx + 'stream'.Length
            if ($ds + 1 -lt $body.Length -and $body.Substring($ds, 2) -eq "`r`n") { $ds += 2 }
            elseif ($ds -lt $body.Length -and $body[$ds] -eq "`n") { $ds += 1 }
            $streamLen = -1
            if ($header -match '/Length\s+(\d+)(?!\s+\d+\s+R)') { $streamLen = [int]$Matches[1] }
            if ($streamLen -lt 0 -and $header -match '/Length\s+(\d+)\s+\d+\s+R') {
                $lenObj = $objects[[int]$Matches[1]]
                if ($lenObj -and -not $lenObj.IsStream) {
                    $lm = [regex]::Match($lenObj.Header, '(\d+)')
                    if ($lm.Success) { $streamLen = [int]$lm.Groups[1].Value }
                }
            }
            $absDataStart = $bodyStart + $ds
            if ($streamLen -ge 0 -and ($absDataStart + $streamLen) -le $text.Length) {
                $endRaw = $absDataStart + $streamLen
            } else {
                $es = $text.IndexOf('endstream', $absDataStart)
                $endRaw = if ($es -ge 0) { $es } else { $endIdx }
            }
            $len = $endRaw - $absDataStart
            if ($len -gt 0) {
                $rawStream = [System.Text.Encoding]::GetEncoding(28591).GetBytes(
                    $text.Substring($absDataStart, $len))
            } else {
                $rawStream = [byte[]]@()
            }
            if ($header -match '/Filter\s*(\[[^\]]*\]|/\w+)') { $filters = $Matches[1] }
            $stream = Expand-PdfFlate $rawStream
            $pos = $endIdx + 6
        } else {
            $pos = $endIdx + 6
        }

        $objects[$objNum] = @{
            Header    = $header
            Stream    = $stream
            RawStream = $rawStream
            Filters   = $filters
            IsStream  = ($null -ne $rawStream)
        }
    }
    return $objects
}

function Get-PdfObject {
    # Returns the object hashtable for an object number, or $null.
    param([hashtable]$Objects, [int]$ObjNum)
    if ($Objects.ContainsKey($ObjNum)) { return $Objects[$ObjNum] }
    return $null
}

function Get-PdfPageList {
    # Returns array of @{ ObjNum; MediaBox; ContentRefs; ResourcesText; ResourcesRef;
    #                     RawHeader } in page-tree order, with MediaBox/Resources
    # inherited from ancestor /Pages nodes when the page lacks them.
    param([hashtable]$Objects)

    $pageSet = @{}
    foreach ($num in $Objects.Keys) {
        if ($Objects[$num].Header -match '/Type\s*/Page(?![sA-Za-z])') { $pageSet[$num] = $true }
    }

    # collect /Pages nodes and their kids
    $kidRefs = @{}
    foreach ($num in $Objects.Keys) {
        $h = $Objects[$num].Header
        if ($h -match '/Type\s*/Pages' -and $h -match '/Kids\s*\[([^\]]*)\]') {
            $kids = @()
            foreach ($mm in [regex]::Matches($Matches[1], '(\d+)\s+\d+\s+R')) {
                $kids += [int]$mm.Groups[1].Value
            }
            $kidRefs[$num] = $kids
        }
    }

    # root = a /Pages node never referenced as someone else's kid
    $rootNum = $null
    foreach ($num in @($kidRefs.Keys)) {
        $isKid = $false
        foreach ($other in @($kidRefs.Keys)) {
            if ($other -ne $num -and $kidRefs[$other] -contains [int]$num) { $isKid = $true; break }
        }
        if (-not $isKid) { $rootNum = [int]$num; break }
    }
    if ($null -eq $rootNum -and $kidRefs.Count -gt 0) {
        $rootNum = [int](@($kidRefs.Keys) | Select-Object -First 1)
    }

    # iterative walk carrying inherited MediaBox/Resources
    $ordered = New-Object System.Collections.ArrayList
    $stack = New-Object System.Collections.Stack
    $visited = @{}
    if ($null -ne $rootNum) { $stack.Push(@([int]$rootNum, $null, $null, $null)) }
    while ($stack.Count -gt 0) {
        $item = $stack.Pop()
        $cur = [int]$item[0]
        $inhMb = $item[1]; $inhResRef = $item[2]; $inhResText = $item[3]
        if ($visited.ContainsKey($cur)) { continue }
        $visited[$cur] = $true
        $o = Get-PdfObject $Objects $cur
        if (-not $o) { continue }
        $h = $o.Header

        if ($pageSet.ContainsKey($cur)) {
            $mb = $null
            if ($h -match '/MediaBox\s*\[\s*([-\d.]+)\s+([-\d.]+)\s+([-\d.]+)\s+([-\d.]+)\s*\]') {
                $mb = @([double]$Matches[1], [double]$Matches[2], [double]$Matches[3], [double]$Matches[4])
            } elseif ($inhMb) { $mb = $inhMb }
            $contentRefs = @()
            if ($h -match '/Contents\s+(\d+)\s+\d+\s+R') { $contentRefs = ,([int]$Matches[1]) }
            elseif ($h -match '/Contents\s*\[([^\]]*)\]') {
                foreach ($mm in [regex]::Matches($Matches[1], '(\d+)\s+\d+\s+R')) {
                    $contentRefs += [int]$mm.Groups[1].Value
                }
            }
            $resText = $null; $resRef = $null
            if ($h -match '/Resources\s+(\d+)\s+\d+\s+R') { $resRef = [int]$Matches[1] }
            elseif ($h -match '/Resources\s*(<<.*?>>)', 'Singleline') { $resText = $Matches[1] }
            if ($null -eq $resRef -and $null -eq $resText) {
                $resRef = $inhResRef; $resText = $inhResText
            }
            [void]$ordered.Add(@{
                ObjNum = $cur
                MediaBox = $mb
                ContentRefs = $contentRefs
                ResourcesText = $resText
                ResourcesRef = $resRef
                RawHeader = $h
            })
            continue
        }
        if ($kidRefs.ContainsKey($cur)) {
            # pass this node's own MediaBox/Resources down to kids
            $downMb = $inhMb
            if ($h -match '/MediaBox\s*\[\s*([-\d.]+)\s+([-\d.]+)\s+([-\d.]+)\s+([-\d.]+)\s*\]') {
                $downMb = @([double]$Matches[1], [double]$Matches[2], [double]$Matches[3], [double]$Matches[4])
            }
            $downRef = $inhResRef; $downText = $inhResText
            if ($h -match '/Resources\s+(\d+)\s+\d+\s+R') { $downRef = [int]$Matches[1]; $downText = $null }
            elseif ($h -match '/Resources\s*(<<.*?>>)', 'Singleline') { $downText = $Matches[1]; $downRef = $null }
            $kids = $kidRefs[$cur]
            for ($k = $kids.Count - 1; $k -ge 0; $k--) {
                [void]$stack.Push(@([int]$kids[$k], $downMb, $downRef, $downText))
            }
        }
    }

    # pages not reached by the walk (broken tree) keep file order
    foreach ($num in @($pageSet.Keys | Sort-Object)) {
        $found = $false
        foreach ($p in $ordered) { if ([int]$p.ObjNum -eq [int]$num) { $found = $true; break } }
        if (-not $found) {
            $o = $Objects[$num]
            $mb = $null
            if ($o.Header -match '/MediaBox\s*\[\s*([-\d.]+)\s+([-\d.]+)\s+([-\d.]+)\s+([-\d.]+)\s*\]') {
                $mb = @([double]$Matches[1], [double]$Matches[2], [double]$Matches[3], [double]$Matches[4])
            }
            [void]$ordered.Add(@{
                ObjNum = [int]$num; MediaBox = $mb; ContentRefs = @()
                ResourcesText = $null; ResourcesRef = $null; RawHeader = $o.Header
            })
        }
    }
    return $ordered.ToArray()
}

function ConvertFrom-PdfCmapText {
    # Parses CMap text into @{ One = hashtable(int->string); Two = hashtable }.
    param([string]$CmapText)
    $one = @{}
    $two = @{}
    # bfchar blocks: <src> <dst>
    foreach ($bm in [regex]::Matches($CmapText, 'beginbfchar(?<body>.*?)endbfchar', 'Singleline')) {
        foreach ($mm in [regex]::Matches($bm.Groups['body'].Value, '<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>')) {
            $src = $mm.Groups[1].Value
            $dstHex = $mm.Groups[2].Value
            $dst = ''
            for ($i = 0; $i + 4 -le $dstHex.Length; $i += 4) {
                $dst += [char][Convert]::ToInt32($dstHex.Substring($i, 4), 16)
            }
            if ($src.Length -le 2) {
                $one[[Convert]::ToInt32($src, 16)] = $dst
            } elseif ($src.Length -le 4) {
                $two[[Convert]::ToInt32($src, 16)] = $dst
            }
        }
    }
    # bfrange blocks: <lo> <hi> <dst>   or   <lo> <hi> [<d1> <d2> ...]
    foreach ($bm in [regex]::Matches($CmapText, 'beginbfrange(?<body>.*?)endbfrange', 'Singleline')) {
        $body = $bm.Groups['body'].Value
        foreach ($mm in [regex]::Matches($body, '<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>\s*(<[0-9A-Fa-f]+>|\[[^\]]*\])')) {
            $lo = [Convert]::ToInt32($mm.Groups[1].Value, 16)
            $hi = [Convert]::ToInt32($mm.Groups[2].Value, 16)
            $dstSpec = $mm.Groups[3].Value
            if ($hi -lt $lo -or ($hi - $lo) -gt 65535) { continue }
            if ($dstSpec -match '^<([0-9A-Fa-f]+)>$') {
                $baseHex = $Matches[1]
                $baseLen = $baseHex.Length
                $base = [Convert]::ToInt32($baseHex, 16)
                for ($c = $lo; $c -le $hi; $c++) {
                    $hex = '{0:X' + [Math]::Max(4, $baseLen) + '}' -f ($base + ($c - $lo))
                    $dst = [string][char][Convert]::ToInt32($hex, 16)
                    if ($baseLen -le 2) { $one[$c] = $dst } else { $two[$c] = $dst }
                }
            } elseif ($dstSpec -match '^\[([^\]]*)\]$') {
                $vals = @()
                foreach ($vm in [regex]::Matches($Matches[1], '<([0-9A-Fa-f]+)>')) {
                    $dst = ''
                    $dstHex = $vm.Groups[1].Value
                    for ($i = 0; $i + 4 -le $dstHex.Length; $i += 4) {
                        $dst += [char][Convert]::ToInt32($dstHex.Substring($i, 4), 16)
                    }
                    $vals += ,$dst
                }
                for ($c = $lo; $c -le $hi -and ($c - $lo) -lt $vals.Count; $c++) {
                    if ($lo -le 255) { $one[$c] = $vals[$c - $lo] } else { $two[$c] = $vals[$c - $lo] }
                }
            }
        }
    }
    return @{ One = $one; Two = $two }
}

function Get-PdfToUnicodeMap {
    # Builds ONE merged byte->string map from all ToUnicode CMaps in the file.
    # Merging is unsafe when fonts reuse the same codes; prefer
    # Get-PdfPageFontMaps (per-font maps selected by Tf). Returns
    # @{ OneByte; TwoByte }.
    param([hashtable]$Objects)

    $one = @{}
    $two = @{}
    foreach ($num in $Objects.Keys) {
        $o = $Objects[$num]
        if (-not $o.IsStream -or $null -eq $o.Stream -or $o.Stream.Length -eq 0) { continue }
        # the CMap operators live in the DECOMPRESSED stream, not in the dict
        $cmapText = [System.Text.Encoding]::GetEncoding(28591).GetString([byte[]]$o.Stream)
        if ($cmapText -notmatch 'beginbfchar' -and $cmapText -notmatch 'beginbfrange') { continue }
        $m = ConvertFrom-PdfCmapText $cmapText
        foreach ($k in $m.One.Keys) { if (-not $one.ContainsKey($k)) { $one[$k] = $m.One[$k] } }
        foreach ($k in $m.Two.Keys) { if (-not $two.ContainsKey($k)) { $two[$k] = $m.Two[$k] } }
    }
    return @{ OneByte = $one; TwoByte = $two }
}

function Get-PdfPageFontMaps {
    # Per-font ToUnicode maps for one page: font resource name (/F1) ->
    # @{ One; Two }. Avoids code-space collisions between fonts.
    param([hashtable]$Objects, $Page)
    $maps = @{}
    $resText = $null
    if ($null -ne $Page.ResourcesText) { $resText = $Page.ResourcesText }
    elseif ($null -ne $Page.ResourcesRef) {
        $ro = Get-PdfObject $Objects ([int]$Page.ResourcesRef)
        if ($ro) { $resText = $ro.Header }
    }
    if (-not $resText) { return $maps }
    $fontDictText = $null
    if ($resText -match '/Font\s+(\d+)\s+\d+\s+R') {
        $fo = Get-PdfObject $Objects ([int]$Matches[1])
        if ($fo) { $fontDictText = $fo.Header }
    } elseif ($resText -match '/Font\s*(<<.*?>>)', 'Singleline') {
        $fontDictText = $Matches[1]
    }
    if (-not $fontDictText) { return $maps }
    foreach ($mm in [regex]::Matches($fontDictText, '/([A-Za-z0-9_.#+\-]+)\s+(\d+)\s+\d+\s+R')) {
        $fname = $mm.Groups[1].Value
        $fontNum = [int]$mm.Groups[2].Value
        if ($maps.ContainsKey($fname)) { continue }
        $fo = Get-PdfObject $Objects $fontNum
        if (-not $fo) { continue }
        if ($fo.Header -match '/ToUnicode\s+(\d+)\s+\d+\s+R') {
            $toObj = Get-PdfObject $Objects ([int]$Matches[1])
            if ($toObj -and $null -ne $toObj.Stream -and $toObj.Stream.Length -gt 0) {
                $cmapText = [System.Text.Encoding]::GetEncoding(28591).GetString([byte[]]$toObj.Stream)
                $maps[$fname] = ConvertFrom-PdfCmapText $cmapText
            }
        }
    }
    return $maps
}

function ConvertFrom-PdfTextBytes {
    # Decodes a PDF string byte sequence using a CMap pair. 2-byte codes win
    # when the map has an entry; otherwise try per-byte; last resort identity.
    param([byte[]]$Bytes, $MapPair)
    if ($null -eq $MapPair) { $MapPair = @{ One = @{}; Two = @{} } }
    $sb = New-Object System.Text.StringBuilder
    $i = 0
    while ($i -lt $Bytes.Length) {
        if ($i + 1 -lt $Bytes.Length) {
            $code2 = [int]$Bytes[$i] * 256 + [int]$Bytes[$i + 1]
            if ($MapPair.Two.ContainsKey($code2)) {
                [void]$sb.Append($MapPair.Two[$code2]); $i += 2; continue
            }
        }
        $b = [int]$Bytes[$i]
        if ($MapPair.One.ContainsKey($b)) {
            [void]$sb.Append($MapPair.One[$b]); $i += 1; continue
        }
        if ($i + 1 -lt $Bytes.Length) {
            # CID fonts without ToUnicode: for CJK, codes often equal Unicode
            if ($code2 -ge 32 -and $code2 -le 0xFFFD) {
                $ch = [string][char]$code2
                [void]$sb.Append($ch); $i += 2; continue
            }
        }
        if ($b -ge 32) { [void]$sb.Append([char]$b) }
        $i += 1
    }
    return $sb.ToString()
}

# merged-map compatibility wrapper (old callers)
function ConvertFrom-PdfTextBytesMerged {
    param([byte[]]$Bytes, [hashtable]$Maps)
    return (ConvertFrom-PdfTextBytes -Bytes $Bytes -MapPair @{ One = $Maps.OneByte; Two = $Maps.TwoByte })
}

function Get-PdfTextWidth {
    # Rough advance-width estimate used for line grouping and column gaps.
    param([string]$Text, [double]$Size)
    $w = 0.0
    foreach ($ch in $Text.ToCharArray()) {
        if ([int]$ch -gt 255) { $w += $Size } else { $w += $Size * 0.5 }
    }
    return $w
}

function ConvertFrom-PdfContentStream {
    # Parses a decompressed content stream; returns text segments:
    # @( @{ X=double; Y=double; Size=double; Text=string } )
    # $FontMaps: optional @{ '<fontName>' = @{ One; Two } } selected by Tf.
    param([byte[]]$StreamBytes, $Maps, $FontMaps = $null)

    if ($null -eq $StreamBytes -or $StreamBytes.Length -eq 0) { return @() }
    if ($null -eq $FontMaps) {
        $FontMaps = @{}
    }
    $s = [System.Text.Encoding]::GetEncoding(28591).GetString($StreamBytes)
    $segments = New-Object System.Collections.ArrayList
    $cur = New-Object System.Text.StringBuilder
    $curBytes = New-Object System.Collections.ArrayList
    $numBuf = ''
    $nums = New-Object System.Collections.ArrayList

    $lineX = 0.0; $lineY = 0.0
    $x = 0.0; $y = 0.0
    $fontSize = 10.0; $leading = 0.0
    $curFont = ''

    $flush = {
        param($Cur, $CurBytes, $PX, $PY, $Size, $Segs)
        if ($CurBytes.Count -gt 0) {
            $pair = $null
            if ($curFont -and $FontMaps.ContainsKey($curFont)) { $pair = $FontMaps[$curFont] }
            elseif ($Maps -and $Maps.OneByte) { $pair = @{ One = $Maps.OneByte; Two = $Maps.TwoByte } }
            if ($pair) {
                $txt = ConvertFrom-PdfTextBytes -Bytes ([byte[]]$CurBytes.ToArray()) -MapPair $pair
            } else {
                $txt = $Cur.ToString()
            }
        } else {
            $txt = $Cur.ToString()
        }
        if ($txt.Length -gt 0) {
            [void]$Segs.Add(@{ X = $PX; Y = $PY; Size = $Size; Text = $txt })
        }
        [void]$Cur.Clear(); [void]$CurBytes.Clear()
    }

    $i = 0; $n = $s.Length
    while ($i -lt $n) {
        $c = $s[$i]
        if ($c -eq '(') {
            $depth = 1; $i++
            while ($i -lt $n -and $depth -gt 0) {
                $ch = $s[$i]
                if ($ch -eq '\') {
                    $i++
                    if ($i -ge $n) { break }
                    $e = $s[$i]
                    if ($e -eq 'n') { [void]$cur.Append("`n") }
                    elseif ($e -eq 't') { [void]$cur.Append("`t") }
                    elseif ($e -eq '(') { [void]$cur.Append('(') }
                    elseif ($e -eq ')') { [void]$cur.Append(')') }
                    elseif ($e -eq '\') { [void]$cur.Append('\') }
                    elseif ($e -match '[0-7]') {
                        $oct = '' + $e
                        while ($oct.Length -lt 3 -and $i + 1 -lt $n -and $s[$i+1] -match '[0-7]') {
                            $i++; $oct += $s[$i]
                        }
                        [void]$curBytes.Add([Convert]::ToInt32($oct, 8) -band 0xFF)
                    }
                    # r/b/f and unknown escapes: skip (CR handled as line break noise)
                } elseif ($ch -eq '(') { $depth++; [void]$cur.Append('(') }
                elseif ($ch -eq ')') { $depth--; if ($depth -gt 0) { [void]$cur.Append(')') } }
                else { [void]$cur.Append($ch) }
                $i++
            }
            $i++; continue
        }
        if ($c -eq '<') {
            if ($i + 1 -lt $n -and $s[$i+1] -eq '<') { $i += 2; continue }
            $i++
            $hex = ''
            while ($i -lt $n -and $s[$i] -ne '>') {
                if ($s[$i] -notmatch '\s') { $hex += $s[$i] }
                $i++
            }
            $i++
            if ($hex.Length % 2 -eq 1) { $hex += '0' }
            for ($h = 0; $h + 1 -lt $hex.Length; $h += 2) {
                [void]$curBytes.Add([Convert]::ToInt32($hex.Substring($h, 2), 16))
            }
            continue
        }
        if ($c -match '[\d.+-]') {
            $numBuf += $c; $i++; continue
        }
        if ($numBuf) {
            $v = 0.0
            if ([double]::TryParse($numBuf, [System.Globalization.NumberStyles]::Float,
                    [System.Globalization.CultureInfo]::InvariantCulture, [ref]$v)) {
                [void]$nums.Add($v)
            }
            $numBuf = ''
        }
        if ($c -eq '/') {
            $i++
            $nameStart = $i
            while ($i -lt $n -and $s[$i] -notmatch '[\s/\[\]<>()]') { $i++ }
            $lastName = $s.Substring($nameStart, $i - $nameStart)
            continue
        }
        if ($c -match '[A-Za-z\*'']') {
            $op = ''
            while ($i -lt $n -and $s[$i] -match '[A-Za-z0-9\*'']') { $op += $s[$i]; $i++ }
            if ($op -eq 'Tj' -or $op -eq "'" -or $op -eq 'TJ') {
                if ($op -eq "'") { $lineY -= $leading; $y = $lineY; $x = $lineX }
                $w = Get-PdfTextWidth $cur.ToString() $fontSize
                & $flush $cur $curBytes $x $y $fontSize $segments
                $x += $w
            }
            elseif ($op -eq 'Tf') {
                if ($nums.Count -ge 1) { $fontSize = [double]$nums[$nums.Count - 1] }
                if ($fontSize -eq 0) { $fontSize = 10 }
                if ($lastName) { $curFont = $lastName }
                $lastName = ''
            }
            elseif ($op -eq 'Td' -and $nums.Count -ge 2) {
                $lineX += [double]$nums[$nums.Count - 2]
                $lineY += [double]$nums[$nums.Count - 1]
                $x = $lineX; $y = $lineY
            }
            elseif ($op -eq 'TD' -and $nums.Count -ge 2) {
                $leading = -1 * [double]$nums[$nums.Count - 1]
                $lineX += [double]$nums[$nums.Count - 2]
                $lineY += [double]$nums[$nums.Count - 1]
                $x = $lineX; $y = $lineY
            }
            elseif ($op -eq 'T*') {
                $lineY -= $leading; $lineX = 0; $x = 0; $y = $lineY
            }
            elseif ($op -eq 'TL' -and $nums.Count -ge 1) { $leading = [double]$nums[$nums.Count - 1] }
            elseif ($op -eq 'Tm' -and $nums.Count -ge 6) {
                $lineX = [double]$nums[$nums.Count - 2]
                $lineY = [double]$nums[$nums.Count - 1]
                $x = $lineX; $y = $lineY
            }
            elseif ($op -eq 'BT') { $lineX = 0; $lineY = 0; $x = 0; $y = 0 }
            [void]$nums.Clear()
            continue
        }
        $i++
    }
    & $flush $cur $curBytes $x $y $fontSize $segments
    return $segments.ToArray()
}

function Get-PdfPageText {
    # Assembles text segments from content streams into visual lines (strings).
    # $FontMaps: per-font CMaps from Get-PdfPageFontMaps (optional; falls back
    # to the merged $Maps).
    param([byte[][]]$ContentStreams, $Maps, $FontMaps = $null)

    $segments = New-Object System.Collections.ArrayList
    foreach ($cs in $ContentStreams) {
        foreach ($seg in (ConvertFrom-PdfContentStream -StreamBytes $cs -Maps $Maps -FontMaps $FontMaps)) {
            [void]$segments.Add($seg)
        }
    }
    if ($segments.Count -eq 0) { return @() }
    $sorted = $segments | Sort-Object -Property `
        @{Expression={ [math]::Round($_.Y, 1) }; Descending=$true },
        @{Expression={ $_.X }; Descending=$false }
    $lines = New-Object System.Collections.ArrayList
    $curLine = $null; $curY = $null; $curEndX = 0.0
    foreach ($seg in $sorted) {
        $tol = [Math]::Max(2.0, [Math]::Abs($seg.Size) * 0.6)
        if ($null -eq $curY -or [Math]::Abs($seg.Y - $curY) -gt $tol) {
            if ($null -ne $curLine) { [void]$lines.Add($curLine) }
            $curLine = $seg.Text
            $curY = $seg.Y
            $curEndX = $seg.X + (Get-PdfTextWidth $seg.Text $seg.Size)
        } else {
            $gap = $seg.X - $curEndX
            if ($gap -gt [Math]::Max(1.0, $seg.Size * 0.18)) { $curLine += ' ' + $seg.Text }
            else { $curLine += $seg.Text }
            $curEndX = $seg.X + (Get-PdfTextWidth $seg.Text $seg.Size)
        }
    }
    if ($null -ne $curLine) { [void]$lines.Add($curLine) }
    return $lines.ToArray()
}

function Get-PdfPageContents {
    # Resolves a page's content streams to decompressed byte arrays.
    param([hashtable]$Objects, $Page)
    $streams = @()
    foreach ($ref in $Page.ContentRefs) {
        $o = Get-PdfObject $Objects ([int]$ref)
        if ($o -and $null -ne $o.Stream -and $o.Stream.Length -gt 0) { $streams += ,([byte[]]$o.Stream) }
    }
    return $streams
}
