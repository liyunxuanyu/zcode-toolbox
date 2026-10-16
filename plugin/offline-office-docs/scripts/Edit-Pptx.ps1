# Edit-Pptx.ps1 - Edit PowerPoint .pptx without Office installed (OOXML via .NET).
# Usage: powershell -NoProfile -ExecutionPolicy Bypass -File Edit-Pptx.ps1 -Action <action> ...
# Actions: gettext | listslides | replace | setslide | addslide

param(
    [Parameter(Mandatory = $true)][string]$Action,
    [string]$Path = '',
    [int]$Slide = 1,
    [string]$Find = '',
    [string]$ReplaceWith = '',
    [string]$TextFile = '',
    [string]$Text = '',
    [string]$Layout = 'titlecontent',
    [string]$OutputPath = ''
)

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot '_Common.ps1')

$NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
$NS_P = 'http://schemas.openxmlformats.org/presentationml/2006/main'

function Get-SlideNames {
    param([string]$ZipPath)
    return @(Get-ZipEntryNames -ZipPath $ZipPath |
        Where-Object { $_ -match '^ppt/slides/slide\d+\.xml$' } |
        Sort-Object { [int]([regex]::Match($_, 'slide(\d+)\.xml').Groups[1].Value) })
}

function Resolve-SlidePart {
    # Returns the zip part name for a 1-based slide number in presentation order.
    param([string]$ZipPath, [int]$SlideNumber)
    $names = @(Get-SlideNames $ZipPath)
    if (@($names).Count -eq 0) { throw 'Presentation has no slides' }
    if ($SlideNumber -lt 1 -or $SlideNumber -gt @($names).Count) {
        throw "Slide number $SlideNumber out of range 1..$(@($names).Count)"
    }
    # presentation.xml sldIdLst defines the real order; slideN.xml naming usually
    # matches, but resolve through rels to be safe.
    $presXml = Read-ZipEntryText -ZipPath $ZipPath -EntryName 'ppt/presentation.xml'
    $pres = New-XmlDoc $presXml
    $pnsm = New-Object System.Xml.XmlNamespaceManager($pres.NameTable)
    $pnsm.AddNamespace('p', $NS_P)
    $pnsm.AddNamespace('r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
    $sldIds = $pres.SelectNodes('//p:sldIdLst/p:sldId', $pnsm)
    if ($sldIds.Count -ge $SlideNumber) {
        $rid = $sldIds[$SlideNumber - 1].GetAttribute('id', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
        $relsXml = Read-ZipEntryText -ZipPath $ZipPath -EntryName 'ppt/_rels/presentation.xml.rels'
        if ($relsXml -and $rid) {
            $m = [regex]::Match($relsXml, "<Relationship [^>]*Id=`"$rid`"[^>]*Target=`"([^`"]+)`"")
            if ($m.Success) {
                $t = $m.Groups[1].Value
                if ($t -notmatch '^/') { $t = 'ppt/' + $t -replace '^ppt/slides/\.\./', 'ppt/' }
                else { $t = $t.TrimStart('/') }
                if (Test-ZipEntryExists -ZipPath $ZipPath -EntryName $t) { return $t }
            }
        }
    }
    return $names[$SlideNumber - 1]
}

function Get-SlideTextLines {
    param([string]$SlideXml)
    $doc = New-XmlDoc $SlideXml
    $nsm = New-Object System.Xml.XmlNamespaceManager($doc.NameTable)
    $nsm.AddNamespace('a', $NS_A)
    $nsm.AddNamespace('p', $NS_P)
    $lines = @()
    foreach ($sp in $doc.SelectNodes('//p:sp', $nsm)) {
        $ph = $sp.SelectSingleNode('.//p:nvSpPr//p:ph/@type', $nsm)
        $prefix = if ($ph) { "[$($ph.Value)] " } else { '' }
        foreach ($para in $sp.SelectNodes('.//a:p', $nsm)) {
            $sb = New-Object System.Text.StringBuilder
            foreach ($t in $para.SelectNodes('.//a:t', $nsm)) { [void]$sb.Append($t.InnerText) }
            $line = $sb.ToString()
            if ($line.Trim().Length -gt 0) { $lines += ($prefix + $line) }
        }
    }
    return $lines
}

switch ($Action.ToLowerInvariant()) {

    'gettext' {
        Assert-InputFile $Path 'Path'
        $names = @(Get-SlideNames $Path)
        $out = New-Object System.Collections.ArrayList
        $i = 0
        foreach ($n in $names) {
            $i++
            [void]$out.Add("=== Slide $i ===")
            foreach ($line in (Get-SlideTextLines (Read-ZipEntryText -ZipPath $Path -EntryName $n))) {
                [void]$out.Add($line)
            }
        }
        $result = $out.ToArray() -join "`n"
        if ($OutputPath) {
            Write-FileTextUtf8 -Path $OutputPath -Text $result
            Write-Info "Text written to: $OutputPath"
        } else {
            Write-Info $result
        }
    }

    'listslides' {
        Assert-InputFile $Path 'Path'
        $names = @(Get-SlideNames $Path)
        Write-Info "Slides: $(@($names).Count)"
        $i = 0
        foreach ($n in $names) {
            $i++
            $lines = @(Get-SlideTextLines (Read-ZipEntryText -ZipPath $Path -EntryName $n))
            $preview = ($lines -join ' ')
            if ($preview.Length -gt 110) { $preview = $preview.Substring(0, 110) + '...' }
            Write-Info ("  Slide {0}: {1}" -f $i, $preview)
        }
    }

    'replace' {
        # Literal replace across every slide (and notes); keeps run formatting of
        # each run where possible by editing a:t text nodes in place.
        Assert-InputFile $Path 'Path'
        if ('' -eq $Find) { throw 'replace needs -Find' }
        Backup-IfNeeded -Path $Path | Out-Null
        $count = 0
        $touchedParts = 0
        foreach ($part in (Get-SlideNames $Path)) {
            $xml = Read-ZipEntryText -ZipPath $Path -EntryName $part
            if (-not $xml.Contains($Find)) { continue }
            $doc = New-XmlDoc $xml
            $nsm = New-Object System.Xml.XmlNamespaceManager($doc.NameTable)
            $nsm.AddNamespace('a', $NS_A)
            $partCount = 0
            foreach ($t in $doc.SelectNodes('//a:t', $nsm)) {
                if ($t.InnerText.Contains($Find)) {
                    $partCount += ([regex]::Matches($t.InnerText, [regex]::Escape($Find))).Count
                    $t.InnerText = $t.InnerText.Replace($Find, $ReplaceWith)
                }
            }
            if ($partCount -gt 0) {
                [void](Write-ZipEntryText -ZipPath $Path -EntryName $part -Text $doc.OuterXml)
                $count += $partCount
                $touchedParts++
            }
        }
        Write-Info "Replaced $count occurrence(s) in $touchedParts slide(s); backup created"
    }

    'setslide' {
        # Replaces the text body of one slide with the given lines (first line
        # becomes the title placeholder content, rest go into the body placeholder).
        Assert-InputFile $Path 'Path'
        $raw = $Text
        if ($TextFile) {
            Assert-InputFile $TextFile 'TextFile'
            $raw = Read-FileTextUtf8 $TextFile
        }
        if ($null -eq $raw) { throw 'setslide needs -Text or -TextFile' }
        Backup-IfNeeded -Path $Path | Out-Null
        $part = Resolve-SlidePart -ZipPath $Path -SlideNumber $Slide
        $xml = Read-ZipEntryText -ZipPath $Path -EntryName $part
        $doc = New-XmlDoc $xml
        $nsm = New-Object System.Xml.XmlNamespaceManager($doc.NameTable)
        $nsm.AddNamespace('a', $NS_A)
        $nsm.AddNamespace('p', $NS_P)

        $lines = @($raw -replace "`r`n", "`n" -replace "`r", "`n" -split "`n")
        $titleLine = ''
        $bodyLines = @()
        if ($lines.Count -gt 0) { $titleLine = $lines[0] }
        if ($lines.Count -gt 1) { $bodyLines = @($lines[1..($lines.Count - 1)]) }

        # find title placeholder sp and first body sp
        $titleSp = $null; $bodySp = $null
        foreach ($sp in $doc.SelectNodes('//p:sp', $nsm)) {
            $ph = $sp.SelectSingleNode('.//p:nvSpPr//p:ph/@type', $nsm)
            $phType = if ($ph) { $ph.Value } else { '' }
            if (-not $titleSp -and $phType -eq 'title') { $titleSp = $sp; continue }
            if (-not $bodySp -and ($phType -eq 'body' -or $phType -eq 'subTitle' -or $phType -eq 'ctrTitle')) {
                if (-not $titleSp -and $phType -eq 'ctrTitle') { $titleSp = $sp; continue }
                $bodySp = $sp
            }
        }
        if (-not $titleSp) {
            # no placeholders (plain text boxes): first sp = title, second = body
            $sps = @($doc.SelectNodes('//p:sp', $nsm))
            if (@($sps).Count -gt 0) { $titleSp = $sps[0] }
            if (-not $bodySp -and @($sps).Count -gt 1) { $bodySp = $sps[1] }
        }

        if ($titleSp) {
            $txBody = $titleSp.SelectSingleNode('.//p:txBody', $nsm)
            if ($txBody) {
                $paras = @($txBody.SelectNodes('./a:p', $nsm))
                if ($paras.Count -gt 0) {
                    # keep first paragraph's run properties; remove others
                    $keep = $paras[0]
                    foreach ($extra in @($paras | Select-Object -Skip 1)) { [void]$txBody.RemoveChild($extra) }
                    $endParaRPr = $keep.SelectSingleNode('./a:endParaRPr', $nsm)
                    foreach ($child in @($keep.ChildNodes)) {
                        if ($child.LocalName -in @('r', 'br', 'fld')) { [void]$keep.RemoveChild($child) }
                    }
                    if ($titleLine.Length -gt 0) {
                        $r = $doc.CreateElement('a', 'r', $NS_A)
                        $t = $doc.CreateElement('a', 't', $NS_A)
                        $t.InnerText = $titleLine
                        [void]$r.AppendChild($t)
                        if ($endParaRPr) { [void]$keep.InsertBefore($r, $endParaRPr) } else { [void]$keep.AppendChild($r) }
                    }
                }
            }
        }
        if ($bodySp) {
            $txBody = $bodySp.SelectSingleNode('.//p:txBody', $nsm)
            if ($txBody) {
                $paras = @($txBody.SelectNodes('./a:p', $nsm))
                $keepP = if ($paras.Count -gt 0) { $paras[0] } else { $null }
                foreach ($extra in @($paras | Select-Object -Skip 1)) { [void]$txBody.RemoveChild($extra) }
                if ($keepP) {
                    # keepP (first a:p) stays as the formatting template at the end
                    # of txBody; append new paragraphs before it so its endParaRPr
                    # remains the last element (its a:r/a:br children are removed)
                    foreach ($child in @($keepP.ChildNodes)) {
                        if ($child.LocalName -in @('r', 'br', 'fld')) { [void]$keepP.RemoveChild($child) }
                    }
                    foreach ($bl in $bodyLines) {
                        if ($bl -eq '' -and @($bodyLines).Count -gt 1) { continue }
                        $pEl = $doc.CreateElement('a', 'p', $NS_A)
                        if ($bl.Length -gt 0) {
                            $r = $doc.CreateElement('a', 'r', $NS_A)
                            $t = $doc.CreateElement('a', 't', $NS_A)
                            $t.InnerText = $bl
                            [void]$r.AppendChild($t)
                            [void]$pEl.AppendChild($r)
                        }
                        [void]$txBody.InsertBefore($pEl, $keepP)
                    }
                }
            }
        }
        [void](Write-ZipEntryText -ZipPath $Path -EntryName $part -Text $doc.OuterXml)
        Write-Info "Slide $Slide text updated ($part); backup created"
    }

    'addslide' {
        # Duplicates the structure of the last slide (keeps layout refs) and
        # replaces its text; simplest reliable way to add slides offline.
        Assert-InputFile $Path 'Path'
        $raw = $Text
        if ($TextFile) {
            Assert-InputFile $TextFile 'TextFile'
            $raw = Read-FileTextUtf8 $TextFile
        }
        Backup-IfNeeded -Path $Path | Out-Null
        $names = @(Get-SlideNames $Path)
        if (@($names).Count -eq 0) { throw 'Presentation has no slides to clone; create it with PowerPoint/WPS first' }
        $srcPart = $names[$names.Count - 1]
        $srcNum = [int]([regex]::Match($srcPart, 'slide(\d+)\.xml').Groups[1].Value)
        $names2 = Get-ZipEntryNames -ZipPath $Path
        $newNum = 1
        while ($names2 -contains "ppt/slides/slide$newNum.xml") { $newNum++ }
        $dstPart = "ppt/slides/slide$newNum.xml"

        # copy slide xml and its rels (layout ref etc.)
        $srcXml = Read-ZipEntryText -ZipPath $Path -EntryName $srcPart
        [void](Write-ZipEntryText -ZipPath $Path -EntryName $dstPart -Text $srcXml)
        $srcRel = "ppt/slides/_rels/slide$srcNum.xml.rels"
        if (Test-ZipEntryExists -ZipPath $Path -EntryName $srcRel) {
            $relBytes = Read-ZipEntryBytes -ZipPath $Path -EntryName $srcRel
            [void](Write-ZipEntryBytes -ZipPath $Path -EntryName "ppt/slides/_rels/slide$newNum.xml.rels" -Bytes $relBytes)
        }

        # content types
        $ctXml = Read-ZipEntryText -ZipPath $Path -EntryName '[Content_Types].xml'
        if ($ctXml -and $ctXml -notmatch [regex]::Escape("/$dstPart")) {
            $ov = [regex]::Match($ctXml, "<Override PartName=`"/$([regex]::Escape($srcPart))`"[^>]*/>")
            if ($ov.Success) {
                $newOv = $ov.Value -replace [regex]::Escape("/$srcPart"), "/$dstPart"
                $newCt = $ctXml -replace '</Types>', "  $newOv`n</Types>"
                [void](Write-ZipEntryText -ZipPath $Path -EntryName '[Content_Types].xml' -Text $newCt)
            }
        }

        # register in presentation.xml + rels
        $presXml = Read-ZipEntryText -ZipPath $Path -EntryName 'ppt/presentation.xml'
        $maxSldId = 255
        foreach ($m in [regex]::Matches($presXml, '<p:sldId id="(\d+)"')) {
            if ([int]$m.Groups[1].Value -gt $maxSldId) { $maxSldId = [int]$m.Groups[1].Value }
        }
        $newSldId = $maxSldId + 1
        $relsXml = Read-ZipEntryText -ZipPath $Path -EntryName 'ppt/_rels/presentation.xml.rels'
        $maxRid = 0
        foreach ($m in [regex]::Matches($relsXml, 'Id="rId(\d+)"')) {
            if ([int]$m.Groups[1].Value -gt $maxRid) { $maxRid = [int]$m.Groups[1].Value }
        }
        $newRid = "rId$($maxRid + 1)"
        $newPres = $presXml -replace '</p:sldIdLst>', "<p:sldId id=`"$newSldId`" r:id=`"$newRid`"/></p:sldIdLst>"
        [void](Write-ZipEntryText -ZipPath $Path -EntryName 'ppt/presentation.xml' -Text $newPres)
        $relTag = "<Relationship Id=`"$newRid`" Type=`"http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide`" Target=`"slides/slide$newNum.xml`"/>"
        $newRels = $relsXml -replace '</Relationships>', "  $relTag`n</Relationships>"
        [void](Write-ZipEntryText -ZipPath $Path -EntryName 'ppt/_rels/presentation.xml.rels' -Text $newRels)

        # set the new slide's text if provided
        if ($raw) {
            & $PSCommandPath -Action setslide -Path $Path -Slide (@($names).Count + 1) -Text $raw
        }
        Write-Info "Added slide ($dstPart) cloned from $srcPart"
    }

    default {
        throw "Unknown action '$Action'. Valid: gettext, listslides, replace, setslide, addslide"
    }
}
