---
description: Convert Office files to PDF (or PDF to Word) via installed MS Office/WPS, offline
argument-hint: "<Office 文件或 PDF 路径> [输出目录]"
allowed-tools: Bash(powershell*)
---

Convert the file(s) "$ARGUMENTS" with the offline-office-docs plugin's COM converter.

Resolve the plugin scripts directory from this command file's location:
`<plugin-root>/scripts/Convert-OfficePdf.ps1` (plugin root is two levels up
from this file's `commands/` folder).

Steps:

1. Parse the arguments: first path is the input (quote it), optional second is
   the output directory. Default output directory = the input file's folder.
2. Run engine detection first:

```
powershell -NoProfile -ExecutionPolicy Bypass -File "<scripts>\Convert-OfficePdf.ps1" -Action detect
```

3. If the input is .doc/.docx/.xls/.xlsx/.ppt/.pptx/.rtf/.csv -> run `-Action to-pdf`
   with `-Path` (single file) or `-Paths` (several), plus `-OutputDir`.
   If the input is .pdf -> run `-Action from-pdf` to produce a .docx.
4. If detection finds no engine, tell the user MS Office or WPS must be installed
   on this machine for conversion, and offer the pure-PowerShell alternatives
   (Edit-Pdf.ps1 createtext / OOXML editing) that need no Office.
5. Verify the output exists and report the produced path(s).
