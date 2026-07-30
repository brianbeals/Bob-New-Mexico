# Prompt for Bob — Document to Data

Paste this with the 16 PDFs attached (or point the tool at this folder).

---

I'm giving you a stack of public New Mexico PDFs: procurement regulations and guidance, quarterly economic summaries, and a couple of health-data documents. They're inconsistent in structure on purpose.

**Do three things.**

**1. Build a structured table, one row per document,** with these columns:
Document Title, Domain, Document Type, Issuing Agency, Effective/Issue Date, Governing Statute or Rule, Key Requirement, Page Count, Text Source (digital vs OCR), and a clickable link that opens the source PDF from this folder.
If a field genuinely isn't in a document, leave it blank. Do not guess or fill it with a plausible value.

**2. Write a 3-to-5 sentence plain-English summary of each document** — what it actually governs, not a restatement of the title. Put these on a second tab.

**3. Tell me the three most interesting things you noticed across the whole set** — especially anything that connects two documents or that a person skimming the pile would miss. Put these on a third tab.

Some files may be scanned images with no text layer. Handle those too (OCR), and flag which ones you had to read that way.

**Deliver it as an Excel workbook (.xlsx), formatted in this house style:**

- **Title banner (row 1):** merged across all columns, dark navy fill `#1E3A5F`, white Arial 13pt bold, centered.
- **Column headers (row 2):** same navy fill `#1E3A5F`, white Arial 10pt bold, left-aligned.
- **Data rows:** alternating fill, white and soft blue `#D6E4F0`; Arial 9pt; first column bold.
- **Borders:** thin gray `#BFBFBF` on every cell.
- **Links:** hyperlink color `#0563C1`, underlined.
- **Sheet setup:** hide gridlines, freeze the header row (freeze panes at A3), zoom 150%.
- Three tabs: Extracted Data, Summaries, Cross-Document Signals.

---

## What a good result looks like
- Missing fields left blank, never fabricated.
- Dates and legal citations correct and verbatim, not approximated.
- The scanned files are read via OCR and flagged, not silently dropped.
- At least one finding connects two separate documents.
