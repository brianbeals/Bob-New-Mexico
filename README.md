# NM Document-to-Data — Demo 1

Sixteen public New Mexico PDFs turned into one clean, queryable table. No template, no data entry. Everything in this repo traces back to those source files.

## What Bob did

- Read all 16 PDFs (11 procurement, 3 economic summaries, 2 public-health data documents)
- Detected four image-only scans (05, 08, 09, 10) with no text layer and ran OCR on them
- Extracted structured metadata for each document: title, domain, type, issuing agency, effective date, governing statute, key requirement, page count
- Wrote a plain-English summary of what each document actually governs
- Identified six cross-document signals — including one that connects a Health Dept data standard to an Economic Development outlier number
- Built a fully formatted Excel workbook matching the house style spec

## Deliverables

| File | What it is |
|------|-----------|
| `index.html` | The published demo page, live at [brianbeals.github.io/Bob-New-Mexico](https://brianbeals.github.io/Bob-New-Mexico/). Carries the same table plus the source-document provenance list and the accessibility before/after |
| `NM-Document-to-Data-Extraction_BB.xlsx` | Main artifact — 3 tabs: Extracted Data, Summaries, Cross-Document Signals |
| `remediated/` | The 16 PDFs after the accessibility pass; originals in the repo root are untouched |
| `NM-Extraction-Data.csv` | Same 16-row table as plain CSV — sortable, filterable, loadable anywhere |
| `Demo-1-Talking-Points.md` | 2–3 minute demo script with the three findings to say out loud |
| `Bob-Prompt-Document-to-Data.md` | The exact prompt used to generate all outputs |
| `build-xlsx.mjs` | Node.js script that regenerates the formatted workbook from scratch |
| `_txt/` | Raw extracted text for all 16 documents (including OCR'd scans) |
| `01-` through `16-*.pdf` | The 16 source PDFs |

## Workbook tabs

**Extracted Data** — one row per document, 11 columns:
`#` · Document Title · Domain · Document Type · Issuing Agency · Effective/Issue Date · Governing Authority · Key Requirement · Pages · Text Source · Open PDF (hyperlinked)

**Summaries** — 3–5 sentence plain-English summary of what each document governs, not a restatement of its title.

**Cross-Document Signals** — six findings worth pointing at in a demo, including the reasoning that connects documents across agencies.

## Formatting spec applied

- Title banner row: dark navy `#1E3A5F`, white Arial 13pt bold, centered
- Column headers: same navy, white Arial 10pt bold
- Data rows: alternating white / soft-blue `#D6E4F0`, Arial 9pt, first two columns bold
- Thin gray `#BFBFBF` borders on every cell
- Hyperlinks: `#0563C1`, underlined, linked to PDFs in this folder
- Gridlines hidden · header row frozen · zoom 150% — all three tabs

## OCR-flagged files

| # | File | Why OCR |
|---|------|---------|
| 05 | `05-Design-Build-Project-Delivery-System.pdf` | Scanned image, no text layer |
| 08 | `08-Proposals-Construction-Facility-Services-Repair.pdf` | Scanned image, no text layer |
| 09 | `09-SPD-Code-of-Ethics.pdf` | Scanned image, no text layer |
| 10 | `10-Government-Advertising-Rates.pdf` | Scanned image, no text layer |

These four are the only fields carrying OCR risk. Everything else is verbatim from digital text.

## Accessibility pass

These are New Mexico state-government records, so ADA Title II applies: 28 CFR Part 35 adopts WCAG 2.1
Level AA for state and local government web content, PDFs included, with New Mexico's compliance date
April 26, 2027 under the 2026 extension. Every file was graded against the machine-detectable gates (a
real text layer, 1.1.1; structure tags, 1.3.1; a declared document language, 3.1.1), then
auto-remediated and re-graded. OCR gave the four image-only scans a text layer, and every file got a
declared language and corrected title metadata: non-compliant went 4 to 0, text layer 12 to 16, declared
language 6 to 16. Structure tagging stayed at 6, because it cannot be done honestly without an authoring
tool, so ten files remain Partial. Alt text, reading order, and contrast still need a person. Before and
after numbers, per document, are on `index.html`.

## To regenerate the workbook

```bash
npm install exceljs
node build-xlsx.mjs
```

## Design choices

- Blank effective-date and statute cells on docs 09 and 15 are intentional — those documents genuinely have neither. Blanks are a feature, not a gap.
- Six signals on the Cross-Document Signals tab; the three strongest are flagged in the talking points. Trim to fit your time.
- Nothing in the outputs names a specific product — kept tool-agnostic per the demo brief.

## One open item

The Procurement Code itself (Chapter 13 NMSA) is a live legal web page, not a PDF, so it is not in the set. Say the word and it can be captured as a 17th document.
