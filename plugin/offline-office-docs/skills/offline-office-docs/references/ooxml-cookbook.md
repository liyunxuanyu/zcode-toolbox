# OOXML cookbook (docx/xlsx/pptx hand edits)

All OOXML editing rests on three helpers in `_Common.ps1`:

- `Read-ZipEntryText -ZipPath <file> -EntryName <part>` — part text as XML string
  (encoding auto-detected from the XML declaration, UTF-8 default).
- `Write-ZipEntryText -ZipPath <file> -EntryName <part> -Text <xml>` — replace part.
- `Test-ZipEntryExists` / `Get-ZipEntryNames` — package discovery.

When a task goes beyond the bundled actions, write a small one-off PowerShell
script that dot-sources `_Common.ps1` and follows the patterns below. Always:

1. `Backup-IfNeeded -Path $Path` before the first write.
2. Re-read the file after writing to verify.

## Namespaces you will need

| Prefix | URI |
|---|---|
| w | http://schemas.openxmlformats.org/wordprocessingml/2006/main |
| s | http://schemas.openxmlformats.org/spreadsheetml/2006/main |
| a | http://schemas.openxmlformats.org/drawingml/2006/main |
| p | http://schemas.openxmlformats.org/presentationml/2006/main |
| r | http://schemas.openxmlformats.org/officeDocument/2006/relationships |
| pr | http://schemas.openxmlformats.org/package/2006/relationships |
| ct | http://schemas.openxmlformats.org/package/2006/content-types |

`Get-XmlNsManager <doc>` pre-registers w/a/s/r/p/cp/dc/dcterms/ct.

## Word (word/document.xml)

Paragraph with a styled run:

```xml
<w:p>
  <w:pPr><w:pStyle w:val="Heading1"/></w:pPr>
  <w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">文本</w:t></w:r>
</w:p>
```

- Every visible character lives in `w:t`. `xml:space="preserve"` keeps
  leading/trailing spaces.
- Bold/italic/font go in `w:rPr` (run properties); alignment/indent in `w:pPr`.
- Body order: paragraphs `w:p`, tables `w:tbl`, then a single trailing `w:sectPr`
  (page setup) — never delete the `sectPr`.
- Styles are defined in `word/styles.xml`; `w:pStyle/@w:val` must match a
  `w:style/@w:styleId` there. To know which style names exist:

```powershell
$styles = Read-ZipEntryText -ZipPath $Path -EntryName 'word/styles.xml'
[regex]::Matches($styles, 'w:styleId="([^"]+)"') | ForEach-Object { $_.Groups[1].Value }
```

- Merged table cells: `w:tcPr` contains `w:gridSpan` (horizontal) or
  `w:vMerge` (vertical; `w:val="restart"` starts, empty `w:vMerge` continues).

## Excel (xl/worksheets/sheetN.xml)

Cell anatomy:

```xml
<c r="B3" t="inlineStr"><is><t xml:space="preserve">文本</t></is></c>
<c r="C3"><v>123.5</v></c>
<c r="D3" t="b"><v>1</v></c>
<c r="E3"><f>SUM(C2:C4)</f></c>           <!-- formula, no cached v -->
<c r="F3" s="4"><v>45000</v></c>          <!-- s = style index -->
```

- `t` attribute: `s`=shared string (value is an index into
  `xl/sharedStrings.xml`), `inlineStr`=literal text, `b`=boolean, `str`=formula
  string result, absent=number.
- Rows are `<row r="3">` inside `<sheetData>`; cells must be in column order
  within a row, rows in ascending order — `ConvertTo-SheetXml` re-sorts for you.
- Update `<dimension ref="A1:F9"/>` when you grow the used range.
- Number formats and fonts live in `xl/styles.xml` (`s="N"` indexes into
  `cellXfs`). Reusing an existing `s` value from a neighbouring cell is the safe
  way to keep formatting — `Set-SheetCell` does this automatically.
- Column widths: `<cols><col min="2" max="2" width="20" customWidth="1"/></cols>`
  before `<sheetData>`.

## PowerPoint (ppt/slides/slideN.xml)

Shape with text:

```xml
<p:sp>
  <p:nvSpPr><p:cNvPr id="5" name="标题 1"/>
    <p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr>
  <p:spPr/>
  <p:txBody>
    <a:bodyPr/><a:lstStyle/>
    <a:p><a:r><a:t>标题文本</a:t></a:r></a:p>
  </p:txBody>
</p:sp>
```

- Placeholder types: `title`, `ctrTitle` (centered, title slide), `body`,
  `subTitle`. Find them with `//p:ph/@type`.
- All text sits in `a:t` inside `a:r` runs inside `a:p` paragraphs.
- Slide order lives in `ppt/presentation.xml` (`p:sldIdLst`); each slide needs
  (a) the part, (b) a relationship in `ppt/_rels/presentation.xml.rels`,
  (c) an override in `[Content_Types].xml`, (d) an entry in `sldIdLst`.
  `Edit-Pptx.ps1 -Action addslide` wires all four by cloning the last slide.

## After any hand edit

Verify by re-reading:

```
powershell -NoProfile -ExecutionPolicy Bypass -File Read-DocInfo.ps1 -Path <file> -FullText
```

and, when available on the machine, `soffice`/Office/WPS opening it is the final
validation. If Office reports the file as corrupt, the usual causes are:
unescaped `<>`& in inserted text (use `ConvertTo-XmlEscapedText`), missing
`xml:space="preserve"`, or a broken `[Content_Types].xml` entry.
