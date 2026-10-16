# Read-DocInfo.ps1 - Structure overview for docx/xlsx/pptx/pdf, plus whole-doc text dump.
# Usage: powershell -NoProfile -ExecutionPolicy Bypass -File Read-DocInfo.ps1 -Path <file> [-FullText] [-Sheet <n>] [-Page <n>]

param(
    [Parameter(Mandatory = $true)][string]$Path,
    [switch]$FullText,
    [int]$Sheet = 0,
    [int]$Page = 0
)

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot '_Common.ps1')
. (Join-Path $PSScriptRoot '_PdfLib.Read.ps1')
. (Join-Path $PSScriptRoot '_PdfLib.Write.ps1')

Assert-InputFile $Path 'Path'
$ext = [System.IO.Path]::GetExtension($Path).ToLowerInvariant()

function Get-OoxmlCoreProps {
    param([string]$ZipPath)
    $core = Read-ZipEntryText -ZipPath $ZipPath -EntryName 'docProps/core.xml'
    $props = @{}
    if ($core) {
        $doc = New-XmlDoc $core
        $nsm = Get-XmlNsManager $doc
        $title = $doc.SelectSingleNode('//dc:title', $nsm)
        $creator = $doc.SelectSingleNode('//dc:creator', $nsm)
        $created = $doc.SelectSingleNode('//dcterms:created', $nsm)
        $modified = $doc.SelectSingleNode('//dcterms:modified', $nsm)
        if ($title) { $props.Title = $title.InnerText }
        if ($creator) { $props.Creator = $creator.InnerText }
        if ($created) { $props.Created = $created.InnerText }
        if ($modified) { $props.Modified = $modified.InnerText }
    }
    $app = Read-ZipEntryText -ZipPath $ZipPath -EntryName 'docProps/app.xml'
    if ($app) {
        $doc = New-XmlDoc $app
        $nsm2 = New-Object System.Xml.XmlNamespaceManager($doc.NameTable)
        $nsm2.AddNamespace('ep', 'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties')
        foreach ($pair in @(@('Application', 'ep:Application'), @('Company', 'ep:Company'))) {
            $n = $doc.SelectSingleNode('//' + $pair[1], $nsm2)
            if ($n) { $props[$pair[0]] = $n.InnerText }
        }
    }
    return $props
}

switch ($ext) {

    '.docx' {
        $props = Get-OoxmlCoreProps $Path
        Write-Info "=== DOCX: $Path ==="
        foreach ($k in $props.Keys) { Write-Info "$k : $($props[$k])" }
        $docXml = Read-ZipEntryText -ZipPath $Path -EntryName 'word/document.xml'
        if (-not $docXml) { throw 'Not a valid .docx (missing word/document.xml)' }
        $doc = New-XmlDoc $docXml
        $nsm = Get-XmlNsManager $doc
        $paras = $doc.SelectNodes('//w:body/w:p', $nsm)
        $tables = $doc.SelectNodes('//w:tbl', $nsm)
        $images = $doc.SelectNodes('//w:drawing', $nsm)
        Write-Info "Paragraphs: $($paras.Count)   Tables: $($tables.Count)   Drawings/Images: $($images.Count)"
        Write-Info ''
        Write-Info '--- First paragraphs (up to 40) ---'
        $i = 0
        foreach ($p in $paras) {
            if ($i -ge 40) { Write-Info '... (use -FullText for everything)'; break }
            $texts = $p.SelectNodes('.//w:t', $nsm)
            $line = ($texts | ForEach-Object { $_.InnerText }) -join ''
            $styleNode = $p.SelectSingleNode('./w:pPr/w:pStyle/@w:val', $nsm)
            $style = if ($styleNode) { " [$($styleNode.Value)]" } else { '' }
            Write-Info ("{0,3}: {1}{2}" -f ($i + 1), $line, $style)
            $i++
        }
        if ($FullText) {
            Write-Info ''
            Write-Info '--- Full text ---'
            Write-Info (Get-OoxmlTextContent $docXml)
        }
    }

    '.xlsx' {
        $props = Get-OoxmlCoreProps $Path
        Write-Info "=== XLSX: $Path ==="
        foreach ($k in $props.Keys) { Write-Info "$k : $($props[$k])" }
        $wbXml = Read-ZipEntryText -ZipPath $Path -EntryName 'xl/workbook.xml'
        if (-not $wbXml) { throw 'Not a valid .xlsx (missing xl/workbook.xml)' }
        $wb = New-XmlDoc $wbXml
        $nsm = Get-XmlNsManager $wb
        $sheets = $wb.SelectNodes('//s:sheet', $nsm)
        Write-Info "Sheets: $($sheets.Count)"
        $i = 0
        foreach ($sh in $sheets) {
            $i++
            $name = $sh.GetAttribute('name')
            $rid = $sh.GetAttribute('id', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
            # resolve r:id -> target via workbook rels
            $target = "xl/worksheets/sheet$i.xml"
            $relsXml = Read-ZipEntryText -ZipPath $Path -EntryName 'xl/_rels/workbook.xml.rels'
            if ($relsXml -and $rid) {
                $rels = New-XmlDoc $relsXml
                $rnsm = New-Object System.Xml.XmlNamespaceManager($rels.NameTable)
                $rnsm.AddNamespace('pr', 'http://schemas.openxmlformats.org/package/2006/relationships')
                $rel = $rels.SelectSingleNode("//pr:Relationship[@Id='$rid']", $rnsm)
                if ($rel) {
                    $t = $rel.GetAttribute('Target')
                    if ($t -notmatch '^/') { $t = 'xl/' + $t }
                    else { $t = $t.TrimStart('/') }
                    $target = $t
                }
            }
            $used = @{ }
            if (Test-ZipEntryExists -ZipPath $Path -EntryName $target) {
                $shXml = Read-ZipEntryText -ZipPath $Path -EntryName $target
                $shDoc = New-XmlDoc $shXml
                $cells = $shDoc.SelectNodes('//s:c', $nsm)
                $rows = $shDoc.SelectNodes('//s:row', $nsm)
                foreach ($c in $cells) {
                    $r = $c.GetAttribute('r')
                    if ($r -match '^([A-Za-z]+)') {
                        $colLetters = $Matches[1]
                        $cn = ConvertTo-ColumnNumber $colLetters
                        if (-not $used.ContainsKey($cn)) { $used[$cn] = 0 }
                        if ($cn -gt $used[$cn]) { $used[$cn] = $cn }
                    }
                }
                $maxCol = 0
                foreach ($k in $used.Keys) { if ($k -gt $maxCol) { $maxCol = $k } }
                $lastColLetter = if ($maxCol -gt 0) { ConvertTo-ColumnLetters $maxCol } else { '-' }
                Write-Info ("  Sheet {0}: '{1}'  rows={2} cells={3} lastCol={4} part={5}" -f $i, $name, $rows.Count, $cells.Count, $lastColLetter, $target)
            } else {
                Write-Info ("  Sheet {0}: '{1}'  (part missing: {2})" -f $i, $name, $target)
            }
        }
        if ($Sheet -gt 0) {
            Write-Info ''
            Write-Info "--- Sheet $Sheet content (CSV, first 200 rows) ---"
            & (Join-Path $PSScriptRoot 'Edit-Xlsx.ps1') -Action export -Path $Path -Sheet $Sheet -OutputPath "$env:TEMP\zz-readinfo-export.csv" 2>$null | Out-Null
            if (Test-FileReadable "$env:TEMP\zz-readinfo-export.csv") {
                $rows = [System.IO.File]::ReadAllLines("$env:TEMP\zz-readinfo-export.csv")
                $shown = 0
                foreach ($r in $rows) {
                    if ($shown -ge 200) { Write-Info '...'; break }
                    Write-Info $r
                    $shown++
                }
                Remove-Item -LiteralPath "$env:TEMP\zz-readinfo-export.csv" -Force -ErrorAction SilentlyContinue
            }
        }
    }

    '.pptx' {
        $props = Get-OoxmlCoreProps $Path
        Write-Info "=== PPTX: $Path ==="
        foreach ($k in $props.Keys) { Write-Info "$k : $($props[$k])" }
        $presXml = Read-ZipEntryText -ZipPath $Path -EntryName 'ppt/presentation.xml'
        if (-not $presXml) { throw 'Not a valid .pptx (missing ppt/presentation.xml)' }
        $slideNames = Get-ZipEntryNames -ZipPath $Path | Where-Object { $_ -match '^ppt/slides/slide\d+\.xml$' } |
            Sort-Object { [int]([regex]::Match($_, 'slide(\d+)\.xml').Groups[1].Value) }
        Write-Info "Slides: $($slideNames.Count)"
        $i = 0
        foreach ($s in $slideNames) {
            $i++
            $sXml = Read-ZipEntryText -ZipPath $Path -EntryName $s
            $sDoc = New-XmlDoc $sXml
            $snsm = Get-XmlNsManager $sDoc
            $texts = $sDoc.SelectNodes('//a:t', $snsm)
            $joined = ($texts | ForEach-Object { $_.InnerText }) -join ' '
            if ($joined.Length -gt 120) { $joined = $joined.Substring(0, 120) + '...' }
            Write-Info ("  Slide {0}: {1}" -f $i, $joined)
            if (-not $FullText -and $i -ge 20) { Write-Info '  ... (use -FullText for all slides)'; break }
        }
        if ($FullText) {
            Write-Info ''
            Write-Info '--- Full slide text ---'
            $j = 0
            foreach ($s in $slideNames) {
                $j++
                Write-Info "=== Slide $j ==="
                Write-Info (Get-OoxmlTextContent (Read-ZipEntryText -ZipPath $Path -EntryName $s))
            }
        }
    }

    '.pdf' {
        Write-Info "=== PDF: $Path ==="
        $bytes = [System.IO.File]::ReadAllBytes($Path)
        $encrypted = -not (Test-PdfNotEncrypted $bytes)
        $sizeKB = [Math]::Round($bytes.Length / 1KB, 1)
        Write-Info "Size: $sizeKB KB"
        Write-Info "Encrypted: $encrypted"
        if ($encrypted) { Write-Warn2 'Encrypted PDF: text extraction unavailable.'; return }
        $objects = Get-PdfRawObjects $bytes
        $maps = Get-PdfToUnicodeMap $objects
        $pages = Get-PdfPageList $objects
        Write-Info "Pages: $($pages.Count)"
        $i = 0
        foreach ($p in $pages) {
            $i++
            $mb = if ($p.MediaBox) { "{0}x{1}mm" -f [Math]::Round(($p.MediaBox[2]-$p.MediaBox[0])*0.3528), [Math]::Round(($p.MediaBox[3]-$p.MediaBox[1])*0.3528) } else { '?' }
            $css = @(Get-PdfPageContents $objects $p)
            $fmaps = Get-PdfPageFontMaps $objects $p
            $lines = @(Get-PdfPageText -ContentStreams $css -Maps $maps -FontMaps $fmaps)
            $preview = ''
            if ($lines.Count -gt 0) {
                $preview = ($lines[0])
                if ($preview.Length -gt 80) { $preview = $preview.Substring(0, 80) + '...' }
            }
            Write-Info ("  Page {0}: {1}, {2} text lines | {3}" -f $i, $mb, $lines.Count, $preview)
            if (-not $FullText -and $i -ge 20) { Write-Info '  ... (use -FullText for all pages)'; break }
        }
        if ($Page -gt 0) {
            if ($Page -gt $pages.Count) { throw "Page $Page out of range (1..$($pages.Count))" }
            Write-Info ''
            Write-Info "--- Page $Page text ---"
            $css = @(Get-PdfPageContents $objects $pages[$Page - 1])
            $fmaps = Get-PdfPageFontMaps $objects $pages[$Page - 1]
            foreach ($line in @(Get-PdfPageText -ContentStreams $css -Maps $maps -FontMaps $fmaps)) { Write-Info $line }
        }
        if ($FullText) {
            Write-Info ''
            Write-Info '--- Full text ---'
            $j = 0
            foreach ($p in $pages) {
                $j++
                Write-Info "=== Page $j ==="
                $css = @(Get-PdfPageContents $objects $p)
                $fmaps = Get-PdfPageFontMaps $objects $p
                foreach ($line in @(Get-PdfPageText -ContentStreams $css -Maps $maps -FontMaps $fmaps)) { Write-Info $line }
            }
        }
    }

    default {
        throw "Unsupported extension '$ext'. Supported: .docx .xlsx .pptx .pdf"
    }
}
