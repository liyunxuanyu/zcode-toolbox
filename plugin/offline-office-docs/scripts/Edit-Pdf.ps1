# Edit-Pdf.ps1 - PDF operations without any external tool.
# Usage: powershell -NoProfile -ExecutionPolicy Bypass -File Edit-Pdf.ps1 -Action <action> ...
# Actions: info | text | merge | split | extract | rotate | createtext

param(
    [Parameter(Mandatory = $true)][string]$Action,
    [string]$Path = '',
    [string[]]$Paths = @(),
    [string]$OutputPath = '',
    [string]$OutputDir = '',
    [string]$Pages = '',
    [int]$Rotate = 90,
    [string]$Text = '',
    [string]$TextFile = '',
    [string]$PageSize = 'A4',
    [double]$FontSize = 11,
    [int]$PerChunk = 1
)

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot '_Common.ps1')
. (Join-Path $PSScriptRoot '_PdfLib.Read.ps1')
. (Join-Path $PSScriptRoot '_PdfLib.Write.ps1')

switch ($Action.ToLowerInvariant()) {

    'info' {
        Assert-InputFile $Path 'Path'
        $bytes = [System.IO.File]::ReadAllBytes($Path)
        $encrypted = -not (Test-PdfNotEncrypted $bytes)
        $count = -1
        $pageSizes = @{}
        if (-not $encrypted) {
            $objects = Get-PdfRawObjects $bytes
            $pageList = @(Get-PdfPageList $objects)
            $count = @($pageList).Count
            foreach ($p in $pageList) {
                if ($p.MediaBox) {
                    $w = [Math]::Round(($p.MediaBox[2] - $p.MediaBox[0]) * 0.3528)
                    $h = [Math]::Round(($p.MediaBox[3] - $p.MediaBox[1]) * 0.3528)
                    $key = "${w}x${h}mm"
                    if ($pageSizes.ContainsKey($key)) { $pageSizes[$key]++ } else { $pageSizes[$key] = 1 }
                }
            }
        }
        $sizeKB = [Math]::Round((Get-Item -LiteralPath $Path).Length / 1KB, 1)
        Write-Info "File: $Path"
        Write-Info "Size: $sizeKB KB"
        Write-Info "Encrypted: $encrypted"
        if ($count -ge 0) {
            Write-Info "Pages: $count"
            $desc = ($pageSizes.GetEnumerator() | Sort-Object Value -Descending |
                ForEach-Object { "$($_.Key) x$($_.Value)" }) -join ', '
            if ($desc) { Write-Info "Page sizes: $desc" }
        } else {
            Write-Warn2 'Encrypted PDF: page count unavailable without decryption.'
        }
    }

    'text' {
        Assert-InputFile $Path 'Path'
        $bytes = [System.IO.File]::ReadAllBytes($Path)
        if (-not (Test-PdfNotEncrypted $bytes)) { throw "PDF is encrypted: '$Path'. Offline toolkit cannot decrypt it." }
        $objects = Get-PdfRawObjects $bytes
        $maps = Get-PdfToUnicodeMap $objects
        # NB: $pageList (not $pages) — names are case-insensitive and $Pages is
        # the page-selection parameter
        $pageList = @(Get-PdfPageList $objects)
        $sel = Resolve-PdfPageSelection $Pages @($pageList).Count
        $out = New-Object System.Collections.ArrayList
        $pn = 0
        foreach ($p in $pageList) {
            $pn++
            if ($sel -notcontains $pn) { continue }
            [void]$out.Add("=== Page $pn ===")
            $css = @(Get-PdfPageContents $objects $p)
            $fmaps = Get-PdfPageFontMaps $objects $p
            foreach ($line in @(Get-PdfPageText -ContentStreams $css -Maps $maps -FontMaps $fmaps)) {
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

    'merge' {
        # -File invocation may deliver "a.pdf","b.pdf" as one string; split on
        # commas that follow a quote or a .pdf/.PDF ending
        $fileList = @()
        foreach ($p in $Paths) {
            $s = ([string]$p).Trim()
            if ($s -match '^["].*["]$') { $s = $s.Trim('"') }
            if ($s -match ',') { $fileList += ($s -split '\s*,\s*' | Where-Object { $_ }) }
            else { $fileList += $s }
        }
        $fileList = @($fileList | Where-Object { $_ })
        if ($fileList.Count -lt 2) { throw 'merge needs -Paths with at least 2 PDF files (comma separated)' }
        foreach ($p in $fileList) { Assert-InputFile $p 'Paths' }
        if (-not $OutputPath) { throw 'merge needs -OutputPath' }
        $made, $count = ConvertTo-PdfMerge -Paths $fileList -OutputPath $OutputPath
        Write-Info "Merged $count pages from $($fileList.Count) files -> $made"
    }

    'split' {
        Assert-InputFile $Path 'Path'
        if ($PerChunk -lt 1) { throw '-PerChunk must be >= 1' }
        $baseName = [System.IO.Path]::GetFileNameWithoutExtension($Path)
        $dir = if ($OutputDir) { $OutputDir } else { (Split-Path -Parent (Resolve-Path -LiteralPath $Path).Path) }
        $made, $total = ConvertTo-PdfSplit -Path $Path -PerChunk $PerChunk -OutputDir $dir -BaseName $baseName
        Write-Info "Split $total pages into $($made.Count) files:"
        foreach ($m in $made) { Write-Info "  $m" }
    }

    'extract' {
        Assert-InputFile $Path 'Path'
        if (-not $OutputPath) { throw 'extract needs -OutputPath' }
        $made, $count = ConvertTo-PdfExtract -Paths @($Path) -Pages $Pages -OutputPath $OutputPath
        Write-Info "Extracted $count page(s) -> $made"
    }

    'rotate' {
        Assert-InputFile $Path 'Path'
        if (-not $OutputPath) { throw 'rotate needs -OutputPath' }
        if ($Rotate -notin @(90, 180, 270, -90)) { throw '-Rotate must be 90, 180, 270 or -90' }
        $made, $total = ConvertTo-PdfRotate -Path $Path -Pages $Pages -Delta $Rotate -OutputPath $OutputPath
        Write-Info "Rotated (selection: $Pages; all pages when empty) -> $made"
    }

    'createtext' {
        $text = $Text
        if ($TextFile) {
            Assert-InputFile $TextFile 'TextFile'
            $text = Read-FileTextUtf8 $TextFile
        }
        if ([string]::IsNullOrWhiteSpace($text)) { throw 'createtext needs -Text or -TextFile with content' }
        if (-not $OutputPath) { throw 'createtext needs -OutputPath' }
        $made, $count = ConvertTo-PdfFromText -Text $text -OutputPath $OutputPath -PageSize $PageSize -FontSize $FontSize
        Write-Info "Created $count page(s) -> $made"
    }

    default {
        throw "Unknown action '$Action'. Valid: info, text, merge, split, extract, rotate, createtext"
    }
}
