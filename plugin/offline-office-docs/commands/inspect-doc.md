---
description: Inspect a Word/Excel/PPT/PDF document's structure offline (docx/xlsx/pptx/pdf)
argument-hint: "<文件路径>"
allowed-tools: Bash(powershell*)
---

Inspect the document at "$ARGUMENTS" using the offline-office-docs plugin.

Resolve the plugin scripts directory from this command file's location:
`<plugin-root>/scripts/Read-DocInfo.ps1` (the plugin root is two levels up from
this file's `commands/` folder).

Run:

```
powershell -NoProfile -ExecutionPolicy Bypass -File "<scripts>\Read-DocInfo.ps1" -Path "$ARGUMENTS"
```

Then report: file type, key counts (paragraphs/tables for docx, sheets for xlsx,
slides for pptx, pages for pdf) and a short content preview. Offer the user the
next logical step (e.g. "export sheet to CSV", "extract pdf text", "replace
text") instead of running everything at once.
