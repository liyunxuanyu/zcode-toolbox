# PDF engine notes and limitations

The bundled engine (`_PdfLib.Read.ps1` / `_PdfLib.Write.ps1`) is a from-scratch
PowerShell PDF parser. Know its boundaries before promising results.

## What works well

- **Text extraction** from text-based PDFs, including Chinese (both simple
  WinAnsi-encoded strings and CID-keyed fonts with `ToUnicode` CMaps — this
  covers most Word/WPS/LaTeX exports).
- Line assembly: text is grouped into visual lines using text-positioning
  operators (`Td`, `TD`, `Tm`, `T*`), wide gaps become spaces.
- **Structure-preserving page operations**: extract / merge / split / rotate.
  Pages are copied with their content streams, resources, annotations, and
  inherited MediaBox/Rotate; xref offsets are recomputed on write, so output is
  well-formed even when the input's xref table is broken (the parser scans
  objects directly instead of trusting the xref).
- **Creating PDFs from plain text** with word wrap; Chinese text uses the
  non-embedded CID font `STSong-Light` (UniGB-UCS2-H), which Acrobat and most
  viewers (including WPS and Edge) render via their Asian font packs.
- Detecting encryption before wasting time on extraction.

## What does NOT work — say so instead of failing silently

| Situation | Behavior / what to tell the user |
|---|---|
| Encrypted / password PDFs | Refused with a clear error; the engine cannot decrypt |
| Scanned or image-only PDFs | No text layer → extraction returns almost nothing; state this explicitly, suggest OCR (out of offline scope) |
| Object streams (PDF 1.5+ compressed xref/objstm) | Page discovery still works (objects are scanned), but text inside objects packed into object streams may be missed; if page count looks wrong, tell the user |
| Fonts without ToUnicode maps | Text decodes via a best-effort CID≈Unicode assumption; garbled CJK means the map was missing — report quality uncertainty |
| Exact layout / columns / tables | Extraction returns reading-order lines, not layout; do not promise pixel-accurate structure |
| PDF forms (AcroForm field values) | Not extracted; only page text is |
| Editing text inside an existing PDF | Not supported — rebuild the PDF (createtext) or edit the source document instead |
| Legacy `.doc`/`.ppt`/`.xls` inputs | Not OOXML; convert via `Convert-OfficePdf.ps1` (needs Office/WPS) first |

## Command examples

```powershell
# page info + first-page preview
powershell -NoProfile -ExecutionPolicy Bypass -File Edit-Pdf.ps1 -Action info -Path "报告.pdf"

# extract text of pages 2-4 to a file (UTF-8)
powershell -NoProfile -ExecutionPolicy Bypass -File Edit-Pdf.ps1 -Action text -Path "报告.pdf" -Pages "2-4" -OutputPath out.txt

# merge (order = argument order)
powershell -NoProfile -ExecutionPolicy Bypass -File Edit-Pdf.ps1 -Action merge -Paths "a.pdf","b.pdf" -OutputPath merged.pdf

# split every 5 pages into part files
powershell -NoProfile -ExecutionPolicy Bypass -File Edit-Pdf.ps1 -Action split -Path big.pdf -PerChunk 5

# extract pages 1,3,7-9 into a new PDF
powershell -NoProfile -ExecutionPolicy Bypass -File Edit-Pdf.ps1 -Action extract -Path big.pdf -Pages "1,3,7-9" -OutputPath sel.pdf

# rotate ALL pages 90° clockwise (or -Pages "2,5")
powershell -NoProfile -ExecutionPolicy Bypass -File Edit-Pdf.ps1 -Action rotate -Path in.pdf -Rotate 90 -OutputPath out.pdf

# create a PDF from text (CJK safe)
powershell -NoProfile -ExecutionPolicy Bypass -File Edit-Pdf.ps1 -Action createtext -TextFile notes.txt -OutputPath notes.pdf -PageSize A4
```

After merge/extract/split/rotate, verify with `-Action info` on the output and
spot-check text with `-Action text -Pages 1`.
