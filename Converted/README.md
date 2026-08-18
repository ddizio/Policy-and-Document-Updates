# Converted Documents

All eight policy/procedure documents from the repository root, reformatted to match
**Procedure and Policy Template Example.docx**.

## What the template look consists of

| Element | Formatting |
|---|---|
| Header block | Bordered table: Vantage logo, document title, `Document Owner:` / `Date:` row |
| Section headings | Full-width green (`#098164`) banner, white bold Aptos 12pt, centred |
| Sub-headings (`3.1`, `4.2`, …) | Century Gothic bold 10pt in green |
| Body text | Century Gothic 10pt |
| Tables | Green header row with white bold text, grey 0.5pt gridlines, repeating header |
| Footer | Page number + document title |
| Page setup | US Letter, 0.5" margins |

## How each document was mapped

The original title block is replaced by the template header. Title, document owner
and date are read out of each source document; where a source had no owner the
header shows `Sustainability`, and where it had no date (or only a `[MM/DD/YYYY]`
placeholder) the header shows the conversion date, 08/18/2026.

Any remaining front-matter metadata (Version, Review Cycle, Approved By, Related
Framework, Reporting Period, Prepared By …) is kept in a **Document Control** table
directly under the header, so nothing from the originals is lost.

Body content — headings, sub-headings, paragraphs, bullet and numbered lists, and
tables — carries over as-is and is restyled. Verified at 98.6–100% word-for-word
retention against the sources; the only intentional removals are the
"Vantage Specialty Chemicals" title line (now the logo), placeholder dates, and
owner/date values that moved into the header.

## Regenerating

```bash
python3 tools/convert_to_template.py            # writes to Converted/
python3 tools/convert_to_template.py outdir     # or a directory of your choice
```

The script reads the template straight from
`Procedure and Policy Template Example.docx`, so editing the template (colours,
fonts, logo) and re-running propagates the change to every document.
