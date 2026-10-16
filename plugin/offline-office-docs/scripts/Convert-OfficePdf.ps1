# Convert-OfficePdf.ps1 - Convert Office files <-> PDF using locally installed
# MS Office or WPS Office via COM automation (fully offline).
# Requires: Word/Excel/PowerPoint or WPS installed on the machine.
# Usage: powershell -NoProfile -ExecutionPolicy Bypass -File Convert-OfficePdf.ps1 -Action <action> ...

param(
    [Parameter(Mandatory = $true)][string]$Action,
    [string]$Path = '',
    [string[]]$Paths = @(),
    [string]$OutputDir = '',
    [string]$OutputPath = '',
    [string]$Engine = 'auto',       # auto | office | wps
    [string]$Format = 'pdf'
)

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot '_Common.ps1')

# --- COM ProgIDs -------------------------------------------------------------
$engines = @{
    word = @(
        @{ ProgId = 'Word.Application';     Name = 'MS Word' },
        @{ ProgId = 'KWPS.Application';     Name = 'WPS Writer' }
    )
    excel = @(
        @{ ProgId = 'Excel.Application';    Name = 'MS Excel' },
        @{ ProgId = 'Ket.Application';      Name = 'WPS Spreadsheets' }
    )
    ppt = @(
        @{ ProgId = 'PowerPoint.Application'; Name = 'MS PowerPoint' },
        @{ ProgId = 'KWPP.Application';       Name = 'WPS Presentation' }
    )
}

function Select-Engine {
    param([string]$Kind)
    $candidates = $engines[$Kind]
    if ($Engine -eq 'office') { $candidates = @($candidates | Where-Object { $_.Name -like 'MS *' }) }
    if ($Engine -eq 'wps')    { $candidates = @($candidates | Where-Object { $_.Name -like 'WPS *' }) }
    foreach ($c in $candidates) {
        try {
            $type = [Type]::GetTypeFromProgID($c.ProgId)
            if ($type) { return $c }
        } catch { }
    }
    throw "No COM automation engine found for $Kind (tried: $(($candidates | ForEach-Object { $_.ProgId }) -join ', ')). Install MS Office or WPS, or use the pure-PowerShell scripts for OOXML editing."
}

function Get-OutFilePath {
    param([string]$SourcePath, [string]$Dir, [string]$Ext)
    $base = [System.IO.Path]::GetFileNameWithoutExtension($SourcePath)
    $dir2 = if ($Dir) { New-OutputDir $Dir } else { (Split-Path -Parent (Resolve-Path -LiteralPath $SourcePath).Path) }
    return (Join-Path $dir2 ("$base.$Ext"))
}

function Convert-WordToPdf {
    param([string]$Src, [string]$Out)
    $engine = Select-Engine 'word'
    $app = $null
    try { $app = New-Object -ComObject $engine.ProgId } catch { throw "Failed to start $($engine.Name): $_" }
    try {
        $app.Visible = $false
        $doc = $app.Documents.Open($Src, $false, $true) # ReadOnly
        # 17 = wdExportFormatPDF; ExportAsFixedFormat is more reliable than SaveAs2
        $doc.ExportAsFixedFormat($Out, 17)
        $doc.Close($false)
        return $Out
    } finally {
        $app.Quit()
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($app)
    }
}

function Convert-ExcelToPdf {
    param([string]$Src, [string]$Out)
    $engine = Select-Engine 'excel'
    $app = $null
    try { $app = New-Object -ComObject $engine.ProgId } catch { throw "Failed to start $($engine.Name): $_" }
    try {
        $app.Visible = $false
        $app.DisplayAlerts = $false
        $wb = $app.Workbooks.Open($Src, 0, $true)
        $wb.ExportAsFixedFormat(0, $Out) # 0 = xlTypePDF
        $wb.Close($false)
        return $Out
    } finally {
        $app.Quit()
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($app)
    }
}

function Convert-PptToPdf {
    param([string]$Src, [string]$Out)
    $engine = Select-Engine 'ppt'
    $app = $null
    try { $app = New-Object -ComObject $engine.ProgId } catch { throw "Failed to start $($engine.Name): $_" }
    try {
        # PowerPoint has no Visible=$false COM path; keep defaults
        $pres = $app.Presentations.Open($Src, $true, $false, $false) # ReadOnly, no window
        $pres.ExportAsFixedFormat($Out, 2) # 2 = ppFixedFormatTypePDF
        $pres.Close()
        return $Out
    } finally {
        $app.Quit()
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($app)
    }
}

function Convert-PdfToWord {
    # Uses Word (2013+) which can open PDFs for conversion. WPS Writer may work too.
    param([string]$Src, [string]$Out)
    $engine = Select-Engine 'word'
    $app = $null
    try { $app = New-Object -ComObject $engine.ProgId } catch { throw "Failed to start $($engine.Name): $_" }
    try {
        $app.Visible = $false
        $doc = $app.Documents.Open($Src, $false, $true)
        $fmt = 16 # wdFormatXMLDocument (.docx)
        $doc.SaveAs([ref]$Out, [ref]$fmt)
        $doc.Close($false)
        return $Out
    } finally {
        $app.Quit()
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($app)
    }
}

function Convert-ToOfficeFormat {
    # PDF -> docx via Word; office->office conversions via SaveAs (e.g. docx->pdf handled above)
    param([string]$Src, [string]$Out, [string]$TargetExt)
    switch ($TargetExt.ToLowerInvariant()) {
        'docx' { Convert-PdfToWord -Src $Src -Out $Out }
        default { throw "Target format '.$TargetExt' not supported by COM converter (use pdf as target)" }
    }
}

function Get-OfficeKind {
    param([string]$File)
    switch ([System.IO.Path]::GetExtension($File).ToLowerInvariant()) {
        '.doc'  { return 'word' }
        '.docx' { return 'word' }
        '.rtf'  { return 'word' }
        '.xls'  { return 'excel' }
        '.xlsx' { return 'excel' }
        '.csv'  { return 'excel' }
        '.ppt'  { return 'ppt' }
        '.pptx' { return 'ppt' }
        default { return $null }
    }
}

switch ($Action.ToLowerInvariant()) {

    'to-pdf' {
        $files = @($Paths)
        if ($files.Count -eq 0 -and $Path) { $files = @($Path) }
        if ($files.Count -eq 0) { throw 'to-pdf needs -Path or -Paths' }
        foreach ($f in $files) {
            Assert-InputFile $f 'Path'
            $kind = Get-OfficeKind $f
            if (-not $kind) { throw "Unsupported input type for to-pdf: '$f'" }
            $out = if ($OutputPath -and $files.Count -eq 1) { $OutputPath } else { Get-OutFilePath $f $OutputDir 'pdf' }
            switch ($kind) {
                'word'  { Convert-WordToPdf -Src $f -Out $out | Out-Null }
                'excel' { Convert-ExcelToPdf -Src $f -Out $out | Out-Null }
                'ppt'   { Convert-PptToPdf -Src $f -Out $out | Out-Null }
            }
            Write-Info "Converted: $f -> $out"
        }
    }

    'from-pdf' {
        Assert-InputFile $Path 'Path'
        if ((Get-Item -LiteralPath $Path).Extension -ne '.pdf') { throw 'from-pdf expects a .pdf input' }
        if (-not $OutputPath) {
            $OutputPath = Get-OutFilePath $Path $OutputDir 'docx'
        }
        Convert-ToOfficeFormat -Src $Path -Out $OutputPath -TargetExt 'docx' | Out-Null
        Write-Info "Converted: $Path -> $OutputPath (PDF layout may not survive conversion perfectly)"
    }

    'detect' {
        foreach ($kind in @('word', 'excel', 'ppt')) {
            $found = @()
            foreach ($c in $engines[$kind]) {
                $t = $null
                try { $t = [Type]::GetTypeFromProgID($c.ProgId) } catch { }
                if ($t) { $found += "$($c.Name) ($($c.ProgId))" }
            }
            if ($found.Count -gt 0) {
                Write-Info "${kind}: $($found -join ', ')"
            } else {
                Write-Warn2 "${kind}: no COM engine available"
            }
        }
    }

    default {
        throw "Unknown action '$Action'. Valid: to-pdf, from-pdf, detect"
    }
}
