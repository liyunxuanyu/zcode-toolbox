---
name: offline-office-docs
description: Offline office document toolkit for Windows intranet hosts with no internet, no npm and no Python. Read and edit Word (.docx), Excel (.xlsx) and PowerPoint (.pptx) packages directly as OOXML, extract text / merge / split / rotate / create PDFs with a built-in pure-PowerShell PDF engine, and convert Office files to PDF (or PDF to Word) through locally installed MS Office or WPS when available. Use when the user asks to read, create, modify, convert, merge or extract from office documents in an offline environment. All scripts run with Windows PowerShell 5.1 (powershell.exe) + .NET only - nothing to install.
---

# Offline Office Documents (Windows 内网离线办公文档处理)

Everything on this host runs **offline**: pure Windows PowerShell 5.1 + .NET.
No npm, no Python, no downloads. OOXML files (.docx/.xlsx/.pptx) are ZIP+XML
packages edited in place via System.IO.Compression. PDFs are parsed/rewritten by
a built-in PowerShell engine.

## Rules

1. **Run, don't rewrite.** Invoke the bundled scripts with powershell.exe; do not
   paste their logic into new ad-hoc scripts.
2. **Read before write.** Always run `Read-DocInfo.ps1` (or the script's read
   action) first so you know sheet names, paragraph counts, page counts.
3. **Backups are automatic** for every modifying action (`*.bak-<timestamp>` next
   to the file). Tell the user when a backup was created.
4. **Chinese text**: pass arguments as UTF-8; CSV import/export uses UTF-8 with
   BOM and falls back to GB18030 when reading. Verify edited files by re-reading
   them afterwards.
5. Quote every path. Paths with spaces must be wrapped in quotes.

## How to invoke

All scripts live in this plugin's `scripts/` directory. Resolve it as
`$PSScriptRoot`-relative from the script you call; in commands below replace
`<scripts>` with the absolute plugin scripts path
(typically `...\plugins\offline-office-docs\<version>\scripts` where you found this SKILL.md).

Pattern:

```
powershell -NoProfile -ExecutionPolicy Bypass -File "<scripts>\Read-DocInfo.ps1" -Path "C:\docs\报告.docx"
```

## Capability map

| Task | Script & action |
|---|---|
| Inspect any docx/xlsx/pptx/pdf structure | `Read-DocInfo.ps1 -Path <file> [-FullText] [-Sheet n] [-Page n]` |
| Word: dump text / paragraphs | `Edit-Docx.ps1 -Action gettext \| listparas` |
| Word: create doc | `Edit-Docx.ps1 -Action newdoc -OutputPath out.docx -TextFile in.txt -Title 标题 -Author 作者` |
| Word: replace body / append / find-replace | `Edit-Docx.ps1 -Action settext \| appendtext \| replace` |
| Excel: sheet names / read cells | `Edit-Xlsx.ps1 -Action listnames \| getcells -Cells A1,B2:D5` |
| Excel: write cells (auto number/text/bool/formula) | `Edit-Xlsx.ps1 -Action setcells -Cells "B3=文本","C4=123","D5==SUM(C2:C4)"` |
| Excel: export sheet to CSV / import CSV / append rows | `Edit-Xlsx.ps1 -Action export \| import \| addrows` |
| Excel: add / delete sheet (optionally from CSV) | `Edit-Xlsx.ps1 -Action newsheet \| delsheet -SheetName 名称 [-Replace]` |
| PPT: list slides / dump text | `Edit-Pptx.ps1 -Action listslides \| gettext` |
| PPT: replace text everywhere / rewrite one slide | `Edit-Pptx.ps1 -Action replace \| setslide -Slide 2` |
| PPT: add slide (clones last) | `Edit-Pptx.ps1 -Action addslide -Text "新页标题`n要点1`n要点2"` |
| PDF: info / extract text (CJK-capable) | `Edit-Pdf.ps1 -Action info \| text [-Pages "1,3,5-8"]` |
| PDF: merge / split / extract pages / rotate | `Edit-Pdf.ps1 -Action merge \| split \| extract \| rotate` |
| PDF: create from text (A4/A3/A5/Letter/Legal, CJK ok) | `Edit-Pdf.ps1 -Action createtext -TextFile in.txt -OutputPath out.pdf` |
| Office -> PDF / PDF -> Word via installed Office or WPS | `Convert-OfficePdf.ps1 -Action to-pdf \| from-pdf \| detect` |

Read `references/ooxml-cookbook.md` before hand-editing document XML beyond the
bundled actions (styles, images, new paragraphs with formatting).
Read `references/pdf-notes.md` for PDF engine limitations (encrypted files,
image-only PDFs, scanned documents) before promising PDF results.

## Excel round-trip recipe (most reliable edit pattern)

For spreadsheets you would rather drive as data:

1. `Edit-Xlsx.ps1 -Action export -Path book.xlsx -Sheet 1 -OutputPath sheet1.csv`
2. Edit the CSV as text (careful with commas; the CSV is RFC4180-quoted).
3. `Edit-Xlsx.ps1 -Action import -Path book.xlsx -Sheet 1 -InputPath sheet1.csv -Replace`
   (or `-Action addrows` to append without replacing).

This preserves number/bool typing; the cell-level `setcells` action is better for
surgical changes (`-CellTypes "B2=text"` to force a type).

## Choosing PDF vs COM conversion

- If the user needs **visual fidelity** (layout, images, fonts) -> `Convert-OfficePdf.ps1`
  (requires Office/WPS on the machine; check with `-Action detect` first).
- If only **text/data** matters, or no Office is installed -> the pure-PowerShell
  PDF engine (`Edit-Pdf.ps1`) and OOXML scripts always work.

## Common pitfalls

- `.xls` / `.ppt` (legacy binary formats) are NOT OOXML: convert them with
  `Convert-OfficePdf.ps1` (COM) or ask the user to save as x/docx first.
- Encrypted / password-protected PDFs cannot be processed; report this clearly.
- Scanned (image-only) PDFs contain no text layer - extraction returns little or
  nothing; say so instead of returning an empty file as success.
- When a script throws "file may be encrypted or damaged" on merge/extract, the
  source PDF likely uses object streams or encryption; try `Edit-Pdf.ps1 -Action info`
  first to confirm what is readable.
- Excel formulas written via `setcells` are stored but not calculated: Excel/WPS
  computes them on open. Cached values shown by other tools may be stale.
