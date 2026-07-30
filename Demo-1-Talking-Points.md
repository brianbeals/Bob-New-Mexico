# Demo 1 — Document to Data (2 to 3 minutes)

## The one-liner to open with
"I took 16 public New Mexico PDFs off state websites, the same pile every agency has on its desk, and turned them into one clean, queryable table. No template, no cleanup, no data entry. Here is what it found."

## What was in the pile (say this fast)
16 documents, three different worlds, on purpose:

- 11 procurement documents from General Services: regulations, guidance memos, a code of ethics, an IT-consolidation FAQ.
- 3 quarterly economic summaries from the Economic Development Department: statewide, Bernalillo County, Catron County.
- 2 public-health data documents from the Department of Health.

The structure is inconsistent from one to the next. That inconsistency is the test, not a problem.

## The two outputs to show
1. **Extracted Data tab** — every document as a row, with the fields you would actually want: title, type, issuing agency, effective date, governing statute, key requirement, page count. This is the "my Tuesday" moment for the room.
2. **Summaries tab** — a short plain-English summary of what each document actually governs, not a restatement of its title.

Then flip to the **Cross-Document Signals tab**. That is where it stops being data entry and starts being reasoning.

## The three findings to say out loud

**1. It left blanks blank.**
Two documents genuinely have no effective date and no legal citation: the Code of Ethics and the Small Numbers Rule. Those cells are empty. It did not invent a date to make the grid look finished. In a government record, a made-up date is worse than a blank one, and it knew the difference.

**2. A quarter of the pile was unreadable, and it read them anyway.**
Four of the 16 were scanned images with no text inside. Copy-paste returns nothing. It detected that, ran OCR, and pulled the effective dates and statute numbers out of the picture. One even had a citation the scan had smeared into "14-1 1-7"; it recovered the real cite, 14-11-7.

**3. It connected two agencies to explain one number.**
Catron County's economy looks like it jumped 34.7% in a quarter. Bernalillo moved 0.04%. Same report template. The Health Department's county classification, a totally separate document, tags Bernalillo "Metropolitan" and Catron "Rural," and the Health Department's Small Numbers Rule explains why tiny populations swing wildly. It linked a health data standard to an economic report to explain the outlier. Nobody flipping through 16 unrelated PDFs makes that jump in the time it takes to read one.

## The line to close on
"The tool is beside the point. This is what document-to-data does for an agency: the pile on the desk becomes a table you can sort, filter, and trust, in minutes instead of a week."

## If someone asks "how do I know it is right"
- Every field traces back to the source PDF; the statute cites and dates are verbatim from the documents.
- Missing fields are left blank, never guessed.
- The scanned files were OCR'd, so the few numbers there carry normal OCR risk and are the ones worth a human glance. That honesty is the point.

## Backup facts if the room gets specific
- Oldest rule still in force: Design-Build, effective April 15, 1998. Newest: the IT-consolidation FAQ, January 31, 2025. A 27-year span, all current.
- "Chapter 4" means two different things in this pile: State Procurement under one NMAC title, Expenditure of Public Funds under another. The table keeps them separate.
- Statewide taxable gross receipts rose 6% year over year in Q1 FY26; retail trade is the largest industry at 19%.
