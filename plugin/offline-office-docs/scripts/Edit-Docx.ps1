# Edit-Docx.ps1 - Edit Word .docx without Word installed (OOXML via .NET).
# Usage: powershell -NoProfile -ExecutionPolicy Bypass -File Edit-Docx.ps1 -Action <action> ...
# Actions: gettext | settext | replace | newdoc | appendtext | listparas

param(
    [Parameter(Mandatory = $true)][string]$Action,
    [string]$Path = '',
    [string]$Template = 'blank',
    [string]$OutputPath = '',
    [string]$TextFile = '',
    [string]$Text = '',
    [string]$Find = '',
    [string]$ReplaceWith = '',
    [string]$Title = '',
    [string]$Author = '',
    [int]$MaxParas = 0
)

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot '_Common.ps1')

$NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

function Get-DocXml {
    param([string]$ZipPath)
    $xml = Read-ZipEntryText -ZipPath $ZipPath -EntryName 'word/document.xml'
    if (-not $xml) { throw 'Not a valid .docx (missing word/document.xml)' }
    return $xml
}

function Get-WDocNsManager {
    param($Doc)
    $nsm = New-Object System.Xml.XmlNamespaceManager($Doc.NameTable)
    $nsm.AddNamespace('w', $NS_W)
    return ,$nsm
}

function Get-ParaText {
    # Concatenates all w:t of a paragraph.
    param($ParaEl, [System.Xml.XmlNamespaceManager]$Nsm)
    $sb = New-Object System.Text.StringBuilder
    foreach ($t in $ParaEl.SelectNodes('.//w:t', $Nsm)) { [void]$sb.Append($t.InnerText) }
    return $sb.ToString()
}

function New-RunXml {
    # Builds a <w:r> with one <w:t>. Returns the element (caller provides doc).
    param($Doc, [System.Xml.XmlNamespaceManager]$Nsm, [string]$Value, $StyleRun = $null)
    $r = $Doc.CreateElement('w', 'r', $NS_W)
    if ($StyleRun) {
        $rprSrc = $StyleRun.SelectSingleNode('./w:rPr', $Nsm)
        if ($rprSrc) { [void]$r.AppendChild($Doc.ImportNode($rprSrc, $true)) }
    }
    $t = $Doc.CreateElement('w', 't', $NS_W)
    # 2-arg form: .NET maps xml:* to the reserved prefix (a 3-arg call with the
    # xml namespace URI generates d5p1: prefixes that LibreOffice/WPS reject)
    [void]$t.SetAttribute('xml:space', 'preserve')
    $t.InnerText = $Value
    [void]$r.AppendChild($t)
    return $r
}

function Split-TextToParas {
    # Splits input text (LF/CRLF) into non-empty-trimmed paragraph strings.
    param([string]$Raw)
    $norm = $Raw -replace "`r`n", "`n" -replace "`r", "`n"
    $parts = $norm -split "`n"
    $out = @()
    foreach ($p in $parts) { $out += $p }
    return $out
}

function Set-CoreProps {
    param([string]$ZipPath, [string]$Title, [string]$Author)
    $core = Read-ZipEntryText -ZipPath $ZipPath -EntryName 'docProps/core.xml'
    if (-not $core) { return }
    $changed = $false
    if ($Title) {
        $core = [regex]::Replace($core, '<dc:title>.*?</dc:title>', '<dc:title>' + (ConvertTo-XmlEscapedText $Title) + '</dc:title>')
        if ($core -notmatch '<dc:title>') {
            $core = $core -replace '</cp:coreProperties>', '<dc:title>' + (ConvertTo-XmlEscapedText $Title) + '</dc:title></cp:coreProperties>'
        }
        $changed = $true
    }
    if ($Author) {
        $core = [regex]::Replace($core, '<dc:creator>.*?</dc:creator>', '<dc:creator>' + (ConvertTo-XmlEscapedText $Author) + '</dc:creator>')
        if ($core -notmatch '<dc:creator>') {
            $core = $core -replace '</cp:coreProperties>', '<dc:creator>' + (ConvertTo-XmlEscapedText $Author) + '</dc:creator></cp:coreProperties>'
        }
        $changed = $true
    }
    if ($changed) {
        [void](Write-ZipEntryText -ZipPath $ZipPath -EntryName 'docProps/core.xml' -Text $core)
    }
}

switch ($Action.ToLowerInvariant()) {

    'gettext' {
        Assert-InputFile $Path 'Path'
        $xml = Get-DocXml $Path
        $text = Get-OoxmlTextContent $xml
        if ($OutputPath) {
            Write-FileTextUtf8 -Path $OutputPath -Text $text
            Write-Info "Text written to: $OutputPath"
        } else {
            Write-Info $text
        }
    }

    'listparas' {
        Assert-InputFile $Path 'Path'
        $doc = New-XmlDoc (Get-DocXml $Path)
        $nsm = Get-WDocNsManager $doc
        $paras = $doc.SelectNodes('//w:body/w:p', $nsm)
        Write-Info "index|text"
        $i = 0
        foreach ($p in $paras) {
            $i++
            if ($MaxParas -gt 0 -and $i -gt $MaxParas) { Write-Info '...'; break }
            $styleNode = $p.SelectSingleNode('./w:pPr/w:pStyle/@w:val', $nsm)
            $style = if ($styleNode) { " [$($styleNode.Value)]" } else { '' }
            $line = Get-ParaText $p $nsm
            if ($line.Length -gt 100) { $line = $line.Substring(0, 100) + '...' }
            Write-Info "$i|$line$style"
        }
        Write-Info "Total paragraphs: $($paras.Count)"
    }

    'settext' {
        # Replaces the whole body with the given text (one paragraph per line).
        Assert-InputFile $Path 'Path'
        $raw = $Text
        if ($TextFile) {
            Assert-InputFile $TextFile 'TextFile'
            $raw = Read-FileTextUtf8 $TextFile
        }
        if ($null -eq $raw) { throw 'settext needs -Text or -TextFile' }
        Backup-IfNeeded -Path $Path | Out-Null
        $doc = New-XmlDoc (Get-DocXml $Path)
        $nsm = Get-WDocNsManager $doc
        $body = $doc.SelectSingleNode('//w:body', $nsm)
        $sectPr = $body.SelectSingleNode('./w:sectPr', $nsm) # keep page setup
        $firstStyle = $null
        $firstPara = $body.SelectSingleNode('./w:p', $nsm)
        if ($firstPara) {
            $firstStyle = $firstPara.SelectSingleNode('./w:pPr', $nsm)
        }

        # remove all paragraphs (and stray tables) but keep sectPr
        foreach ($child in @($body.ChildNodes)) {
            if ($child.LocalName -eq 'sectPr') { continue }
            [void]$body.RemoveChild($child)
        }

        $anchor = $sectPr
        foreach ($line in (Split-TextToParas $raw)) {
            $p = $doc.CreateElement('w', 'p', $NS_W)
            if ($firstStyle) { [void]$p.AppendChild($doc.ImportNode($firstStyle, $true)) }
            if ($line.Length -gt 0) {
                [void]$p.AppendChild((New-RunXml $doc $nsm $line))
            }
            [void]$body.InsertBefore($p, $anchor)
        }
        Set-CoreProps -ZipPath $Path -Title $Title -Author $Author
        [void](Write-ZipEntryText -ZipPath $Path -EntryName 'word/document.xml' -Text $doc.OuterXml)
        Write-Info ("Body replaced with {0} paragraph(s); backup created" -f @(Split-TextToParas $raw).Count)
    }

    'appendtext' {
        Assert-InputFile $Path 'Path'
        $raw = $Text
        if ($TextFile) {
            Assert-InputFile $TextFile 'TextFile'
            $raw = Read-FileTextUtf8 $TextFile
        }
        if ($null -eq $raw) { throw 'appendtext needs -Text or -TextFile' }
        Backup-IfNeeded -Path $Path | Out-Null
        $doc = New-XmlDoc (Get-DocXml $Path)
        $nsm = Get-WDocNsManager $doc
        $body = $doc.SelectSingleNode('//w:body', $nsm)
        $sectPr = $body.SelectSingleNode('./w:sectPr', $nsm)
        foreach ($line in (Split-TextToParas $raw)) {
            $p = $doc.CreateElement('w', 'p', $NS_W)
            if ($line.Length -gt 0) {
                [void]$p.AppendChild((New-RunXml $doc $nsm $line))
            }
            [void]$body.InsertBefore($p, $sectPr)
        }
        [void](Write-ZipEntryText -ZipPath $Path -EntryName 'word/document.xml' -Text $doc.OuterXml)
        Write-Info ("Appended {0} paragraph(s); backup created" -f @(Split-TextToParas $raw).Count)
    }

    'replace' {
        # Literal (non-regex) replace across runs: rebuilds affected paragraphs
        # keeping the first run's formatting.
        Assert-InputFile $Path 'Path'
        if ('' -eq $Find) { throw 'replace needs -Find' }
        Backup-IfNeeded -Path $Path | Out-Null
        $doc = New-XmlDoc (Get-DocXml $Path)
        $nsm = Get-WDocNsManager $doc
        $paras = $doc.SelectNodes('//w:p', $nsm)
        $count = 0
        $touched = 0
        foreach ($p in $paras) {
            $full = Get-ParaText $p $nsm
            if (-not $full.Contains($Find)) { continue }
            $count += ([regex]::Matches($full, [regex]::Escape($Find))).Count
            $touched++
            $newFull = $full.Replace($Find, $ReplaceWith)
            $runs = $p.SelectNodes('./w:r', $nsm)
            $firstRun = if ($runs.Count -gt 0) { $runs[0] } else { $null }
            # remove all runs (and hyperlink content), then append one rebuilt run
            # (the first run was removed too, so it cannot serve as an insert anchor)
            foreach ($child in @($p.ChildNodes)) {
                if ($child.LocalName -in @('r', 'hyperlink', 'smartTag')) { [void]$p.RemoveChild($child) }
            }
            $newRun = New-RunXml $doc $nsm $newFull $firstRun
            [void]$p.AppendChild($newRun)
        }
        [void](Write-ZipEntryText -ZipPath $Path -EntryName 'word/document.xml' -Text $doc.OuterXml)
        Write-Info "Replaced $count occurrence(s) in $touched paragraph(s); backup created"
    }

    'newdoc' {
        # Creates a new .docx from a bundled template, fills text and core props.
        $tplPath = Join-Path (Split-Path -Parent $PSScriptRoot) "templates\docx-$Template.docx"
        if (-not (Test-FileReadable $tplPath)) {
            throw "Template not found: $tplPath (available: blank)"
        }
        if (-not $OutputPath) { throw 'newdoc needs -OutputPath' }
        $dest = New-OoxmlPackageFromTemplate -TemplatePath $tplPath -DestinationPath $OutputPath
        $raw = ''
        if ($TextFile) {
            Assert-InputFile $TextFile 'TextFile'
            $raw = Read-FileTextUtf8 $TextFile
        } elseif ($Text) { $raw = $Text }
        if ($raw) {
            $doc = New-XmlDoc (Get-DocXml $dest)
            $nsm = Get-WDocNsManager $doc
            $body = $doc.SelectSingleNode('//w:body', $nsm)
            $sectPr = $body.SelectSingleNode('./w:sectPr', $nsm)
            foreach ($line in (Split-TextToParas $raw)) {
                $p = $doc.CreateElement('w', 'p', $NS_W)
                if ($line.Length -gt 0) {
                    [void]$p.AppendChild((New-RunXml $doc $nsm $line))
                }
                [void]$body.InsertBefore($p, $sectPr)
            }
            [void](Write-ZipEntryText -ZipPath $dest -EntryName 'word/document.xml' -Text $doc.OuterXml)
        }
        Set-CoreProps -ZipPath $dest -Title $Title -Author $Author
        Write-Info "Created: $dest"
    }

    default {
        throw "Unknown action '$Action'. Valid: gettext, listparas, settext, appendtext, replace, newdoc"
    }
}
