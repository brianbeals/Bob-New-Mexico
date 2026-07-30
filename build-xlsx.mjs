// build-xlsx.mjs — builds the formatted NM Document-to-Data workbook
// Run from workspace root: node "NM Procurement Docs/build-xlsx.mjs"

import ExcelJS from "exceljs";
import path from "path";
import { fileURLToPath } from "url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const OUT = path.join(__dirname, "NM-Document-to-Data-Extraction_BB.xlsx");

// ── Palette ──────────────────────────────────────────────────────────────────
const NAVY   = "FF1E3A5F";
const WHITE  = "FFFFFFFF";
const BLUE1  = "FFD6E4F0";
const BORDER = "FFBFBFBF";
const LINK   = "FF0563C1";

// ── Helper: thin border on every side ────────────────────────────────────────
function thinBorder() {
  const side = { style: "thin", color: { argb: BORDER } };
  return { top: side, bottom: side, left: side, right: side };
}

// ── Helper: apply navy header style ──────────────────────────────────────────
function navyHeader(cell, bold = true, size = 10) {
  cell.fill   = { type: "pattern", pattern: "solid", fgColor: { argb: NAVY } };
  cell.font   = { name: "Arial", size, bold, color: { argb: WHITE } };
  cell.border = thinBorder();
  cell.alignment = { vertical: "middle", wrapText: true, horizontal: "left" };
}

// ── Helper: data row fill (alternating) ──────────────────────────────────────
function dataFill(rowIndex) {
  // rowIndex is 1-based; row 1 = banner, row 2 = header, rows 3+ = data
  return rowIndex % 2 === 1   // odd data rows (3,5,7…) → white; even → blue
    ? { type: "pattern", pattern: "solid", fgColor: { argb: WHITE } }
    : { type: "pattern", pattern: "solid", fgColor: { argb: BLUE1 } };
}

// ── Data ─────────────────────────────────────────────────────────────────────

const EXTRACTED = [
  {
    num: 1,
    title: "Suspension or Debarment of Bidders, Offerors or Contractors (1.4.7 NMAC)",
    domain: "Procurement",
    type: "State regulation (NMAC rule)",
    agency: "General Services Dept - State Purchasing Division",
    date: "August 30, 2013",
    statute: "Sections 13-1-177 through 13-1-180 NMSA 1978",
    req: "Sets the procedure to suspend or debar contractors; debarment runs 3 months to 3 years.",
    pages: 7,
    source: "Digital",
    file: "01-Debarment-Suspension-Regulations-1.4.7-NMAC.pdf",
  },
  {
    num: 2,
    title: "Procurement Code Regulations (1.4.1 NMAC)",
    domain: "Procurement",
    type: "State regulation (NMAC rule)",
    agency: "General Services Dept - State Purchasing Division",
    date: "June 21, 2022",
    statute: "Sections 9-17-5 and 13-1-95 NMSA 1978; implements Procurement Code 13-1-28 to 13-1-199",
    req: "Master implementing rule for the Procurement Code across executive-branch agencies; lists 11 exemptions.",
    pages: 28,
    source: "Digital",
    file: "02-Procurement-Code-Regulations-Ch4-Part1.pdf",
  },
  {
    num: 3,
    title: "Approval of Contracts for the Purchase of Professional Services (1.4.10 NMAC)",
    domain: "Procurement",
    type: "State regulation (NMAC rule)",
    agency: "General Services Dept - Contracts Review Bureau",
    date: "June 21, 2022",
    statute: "Section 6-5-9 NMSA 1978; Procurement Code 13-1-28 to 13-1-199",
    req: "GSD Contracts Review Bureau must review and approve professional-services contracts over $5,000.",
    pages: 5,
    source: "Digital",
    file: "03-Purchase-of-Professional-Services-Ch4-Part10.pdf",
  },
  {
    num: 4,
    title: "Professional Services vs. General Services - Services Determination",
    domain: "Procurement",
    type: "Guidance / determination (non-rule)",
    agency: "State Purchasing Division",
    date: "November 2022",
    statute: "Sections 13-1-76 and 13-1-87 NMSA 1978",
    req: "Distinguishes professional from general services; agencies must get a State Purchasing Agent determination when a service is unlisted or ambiguous.",
    pages: 3,
    source: "Digital",
    file: "04-Professional-Services-Determination-Guidance-Nov2022.pdf",
  },
  {
    num: 5,
    title: "Design and Build Project Delivery Systems (1.5.7 NMAC)",
    domain: "Procurement",
    type: "State regulation (NMAC rule)",
    agency: "General Services Dept - State Purchasing Division",
    date: "April 15, 1998 (recompiled Nov 30, 2001)",
    statute: "Section 13-1-119.1 NMSA 1978",
    req: "Rules for procuring design-build project delivery; binding on using agencies, excludes road and highway work.",
    pages: 6,
    source: "OCR (scanned image)",
    file: "05-Design-Build-Project-Delivery-System.pdf",
  },
  {
    num: 6,
    title: "Best Value Procurement Guide, Part 1 - Developing Request for RFPs",
    domain: "Procurement",
    type: "Guidance / guide (non-rule)",
    agency: "General Services Dept - State Purchasing Division",
    date: "January 2016",
    statute: "",
    req: "How-to guidance for developing best-value RFPs.",
    pages: 58,
    source: "Digital",
    file: "06-Best-Value-Procurement-Guide-Vol1.pdf",
  },
  {
    num: 7,
    title: "Best Value Procurement Guide, Part 2 - The RFP Process",
    domain: "Procurement",
    type: "Guidance / guide (non-rule)",
    agency: "General Services Dept - State Purchasing Division",
    date: "January 2016",
    statute: "",
    req: "How-to guidance for running the RFP process through evaluation and award.",
    pages: 30,
    source: "Digital",
    file: "07-Best-Value-Procurement-Guide-Vol2.pdf",
  },
  {
    num: 8,
    title: "Competitive Sealed Proposals for Construction, Facility Maintenance, Services and Repairs (1.4.8 NMAC)",
    domain: "Procurement",
    type: "State regulation (NMAC rule)",
    agency: "General Services Dept - State Purchasing Division",
    date: "September 30, 2005",
    statute: "Sections 13-1-67 and 13-1-111 NMSA 1978 (2003 amendments)",
    req: "Uniform procedures for using competitive sealed proposals on construction and facility services.",
    pages: 5,
    source: "OCR (scanned image)",
    file: "08-Proposals-Construction-Facility-Services-Repair.pdf",
  },
  {
    num: 9,
    title: "State Purchasing Division Code of Ethics",
    domain: "Procurement",
    type: "Policy statement (non-rule)",
    agency: "General Services Dept - State Purchasing Division",
    date: "",
    statute: "",
    req: "Ethical-conduct standards for SPD employees: conflicts of interest, gifts, and use of confidential information.",
    pages: 1,
    source: "OCR (scanned image)",
    file: "09-SPD-Code-of-Ethics.pdf",
  },
  {
    num: 10,
    title: "Governmental Advertising Rates (1.4.4 NMAC)",
    domain: "Procurement",
    type: "State regulation (NMAC rule)",
    agency: "General Services Dept - State Purchasing Division",
    date: "February 15, 1999",
    statute: "Section 14-11-7 NMSA 1978",
    req: "Authorizes GSD to set newspaper rates for legally required public notices.",
    pages: 1,
    source: "OCR (scanned image)",
    file: "10-Government-Advertising-Rates.pdf",
  },
  {
    num: 11,
    title: "FAQs - DoIT Purchasing Consolidated with State Purchasing Division",
    domain: "Procurement",
    type: "FAQ memo (informational)",
    agency: "General Services Dept (to Cabinet Secretaries); signed Anna Silva, Acting Cabinet Secretary",
    date: "January 31, 2025",
    statute: "",
    req: "Explains the consolidation of Dept of Information Technology purchasing under the State Purchasing Division.",
    pages: 3,
    source: "Digital",
    file: "11-FAQs-Consolidated-IT-Procurements.pdf",
  },
  {
    num: 12,
    title: "Quarterly Economic Summary - New Mexico (statewide)",
    domain: "Economic",
    type: "Economic data report",
    agency: "NM Economic Development Dept - State Data Center (D. Kwon & A. Anaya, Sr. Economists)",
    date: "January 2026 (covers Q1 FY26)",
    statute: "",
    req: "Statewide matched taxable gross receipts rose $612.1M (2.3%) over the prior quarter and $1.52B (6%) year over year; retail trade was the largest industry at 19%.",
    pages: 4,
    source: "Digital",
    file: "12-EDD-Economic-Summary-Statewide-Jan2026.pdf",
  },
  {
    num: 13,
    title: "Quarterly Economic Summary - Bernalillo County",
    domain: "Economic",
    type: "Economic data report",
    agency: "NM Economic Development Dept - State Data Center (D. Kwon & A. Anaya, Sr. Economists)",
    date: "January 2026 (covers Q1 FY26)",
    statute: "",
    req: "Bernalillo County MTGR was $7.04B, up just $3M (0.04%) from the prior quarter; largest county economy in the state.",
    pages: 4,
    source: "Digital",
    file: "13-EDD-Economic-Summary-Bernalillo-County-Jan2026.pdf",
  },
  {
    num: 14,
    title: "Quarterly Economic Summary - Catron County",
    domain: "Economic",
    type: "Economic data report",
    agency: "NM Economic Development Dept - State Data Center (D. Kwon & A. Anaya, Sr. Economists)",
    date: "January 2026 (covers Q1 FY26)",
    statute: "",
    req: "Catron County MTGR rose about $5.5M (34.7%) over the prior quarter, reflecting the sharp swings typical of a very small rural tax base.",
    pages: 4,
    source: "Digital",
    file: "14-EDD-Economic-Summary-Catron-County-Jan2026.pdf",
  },
  {
    num: 15,
    title: "New Mexico Rule for Small Numbers and Public Data Release",
    domain: "Public Health",
    type: "Data standard / methodology",
    agency: "NM Department of Health (NM-IBIS)",
    date: "",
    statute: "",
    req: "Suppress any data cell when the population is under 20 and the count is 1–3; survey estimates need at least 50 respondents to be released.",
    pages: 1,
    source: "Digital",
    file: "15-NMDOH-Small-Numbers-Rule-Data-Standard.pdf",
  },
  {
    num: 16,
    title: "Metropolitan, Small Metro, Mixed Urban/Rural and Rural New Mexico Counties",
    domain: "Public Health",
    type: "Reference classification",
    agency: "NM Department of Health (NM-IBIS)",
    date: "November 2014",
    statute: "",
    req: "Classifies all 33 counties by urbanization. Bernalillo is Metropolitan; Catron is Rural.",
    pages: 1,
    source: "Digital",
    file: "16-NMDOH-Urban-Rural-County-Classification.pdf",
  },
];

const SUMMARIES = [
  { num: 1, short: "Suspension or Debarment (1.4.7 NMAC)", summary: "This State Purchasing Division rule governs how New Mexico suspends or debars contractors from state contracts, excluding professional services. It sets a debarment period of no less than three months and no more than three years, beginning when the contractor receives the final order. Its stated purpose is to protect the state from contractors unable or unwilling to perform, not to punish. It runs on statutory authority in Sections 13-1-177 through 13-1-180 NMSA 1978 and took effect August 30, 2013." },
  { num: 2, short: "Procurement Code Regulations (1.4.1 NMAC)", summary: "This is the master rule implementing New Mexico's Procurement Code for executive-branch agencies. It applies to nearly every state expenditure on tangible property, services, and construction, then carves out eleven exemptions such as highway construction, the judicial and legislative branches, state universities, and charter schools. It was last replaced effective June 21, 2022. It rests on Sections 9-17-5 and 13-1-95 NMSA 1978 and administers the Code at Sections 13-1-28 through 13-1-199." },
  { num: 3, short: "Professional Services Contract Approval (1.4.10 NMAC)", summary: "This rule requires the GSD Contracts Review Bureau to review and approve every professional-services contract that exceeds $5,000, plus amendments, for most state agencies. It lists exempt agencies including the judicial and legislative branches, state educational institutions, and several named authorities. Notably it sits under NMAC Title 1, Public Finance, Chapter 4 — a different Chapter 4 than the procurement rules. It took effect June 21, 2022 under Section 6-5-9 NMSA 1978." },
  { num: 4, short: "Professional vs. General Services Determination (Nov 2022)", summary: "This State Purchasing Division guidance helps agencies tell professional services from general services, a distinction that changes which procurement path applies. It reproduces the statutory definitions at Sections 13-1-76 and 13-1-87 and lays out characteristics of each, then instructs agencies to seek a State Purchasing Agent determination whenever a service is unlisted or arguable. It is guidance, not a rule, issued November 2022. Its header cites the Code span loosely as 13-1-1 to 13-1-99 while the body relies on 13-1-76 and 13-1-87." },
  { num: 5, short: "Design-Build Project Delivery (1.5.7 NMAC)", summary: "This rule governs how state agencies procure design-build project delivery, where design and construction are contracted together. It binds all using agencies except for state road and highway work. It is one of the oldest documents in the set, effective April 15, 1998 and recompiled November 30, 2001, and sits under a different chapter — Title 1 Chapter 5 on public property procurement. It runs on Section 13-1-119.1 NMSA 1978. The file is a scanned image with no text layer, so its content was recovered by OCR." },
  { num: 6, short: "Best Value Procurement Guide, Part 1", summary: "This is a practical how-to guide, not a regulation, on developing best-value requests for proposals. Issued by the State Purchasing Division in January 2016, it walks agencies through defining requirements and building an RFP that weighs value rather than low price alone. At 58 pages it is the longest document in the set. It cites no single governing statute because it is guidance that supports the Code rather than implementing it." },
  { num: 7, short: "Best Value Procurement Guide, Part 2", summary: "The companion to Part 1, this guide covers running the RFP process itself, from evaluation through award. It too is State Purchasing Division guidance from January 2016 rather than a rule. It gives agencies a repeatable method for scoring proposals on best value. Like Part 1 it names no governing authority because it is instructional support material." },
  { num: 8, short: "Competitive Sealed Proposals for Construction (1.4.8 NMAC)", summary: "This rule sets uniform procedures for using competitive sealed proposals on construction, facility maintenance, services, and repairs. Its goal is to deliver high-quality projects on time, safely, and cost-effectively. It applies to all executive-branch agencies and took effect September 30, 2005 under Sections 13-1-67 and 13-1-111 NMSA 1978. The source file is a scan, so its dates and citations were recovered through OCR." },
  { num: 9, short: "State Purchasing Division Code of Ethics", summary: "This one-page statement sets the ethical standards SPD employees must follow, covering conflicts of interest, gifts and gratuities from suppliers, and the misuse of confidential information. It frames ethical lapses as a breach of public trust and lists concrete do-nots. It carries no effective date and cites no statute, so those fields are genuinely blank rather than missed. The file is a scanned image and was read by OCR." },
  { num: 10, short: "Governmental Advertising Rates (1.4.4 NMAC)", summary: "This rule lets the General Services Department set the newspaper rates paid for legal notices that governmental entities are required to publish. It implements Section 14-11-7 NMSA 1978 and took effect February 15, 1999, making it one of the oldest rules in the set. It applies wherever a public entity must publish notices by law or court order. It is a scanned image; OCR recovered the citation, which the scan had rendered as a garbled '14-1 1-7'." },
  { num: 11, short: "FAQs - DoIT Purchasing Consolidation (Jan 2025)", summary: "This is an informational FAQ memo, not a rule, dated January 31, 2025 and addressed to Cabinet Secretaries. It explains the consolidation of Department of Information Technology purchasing under the State Purchasing Division and answers practical questions about the change. It is the most recent procurement document in the set. It carries a signatory and a date but cites no governing statute, so that field is left blank." },
  { num: 12, short: "Quarterly Economic Summary - New Mexico", summary: "This State Data Center report summarizes New Mexico's economy for the first quarter of fiscal 2026. Statewide matched taxable gross receipts rose $612.1M, or 2.3%, over the prior quarter and $1.52B, or 6%, year over year. Retail trade was the largest industry at 19% of receipts, with construction posting the biggest year-over-year gain. It draws on data from the Taxation and Revenue Department, Workforce Solutions, the BEA, and the Census Bureau, and was released January 2026." },
  { num: 13, short: "Quarterly Economic Summary - Bernalillo County", summary: "This is the Bernalillo County edition of the same quarterly template. The county reported $7.04B in matched taxable gross receipts, the largest economy in the state, yet it moved only $3M, or 0.04%, from the prior quarter. That near-flat change reflects how a very large, diversified tax base absorbs quarterly swings. Released January 2026 by the State Data Center for Q1 FY26." },
  { num: 14, short: "Quarterly Economic Summary - Catron County", summary: "This is the Catron County edition of the same template. Catron's receipts rose about $5.5M, or 34.7%, over the prior quarter — a dramatic percentage swing on a very small base of roughly twenty million dollars. The report itself notes the pattern of sharp first-quarter rises typical of small rural counties. Released January 2026 for Q1 FY26, it is a clean example of small-number volatility next to Bernalillo's stability." },
  { num: 15, short: "NM Rule for Small Numbers and Public Data Release", summary: "This Department of Health data standard governs when small counts must be hidden before public release, to protect individual privacy. It requires suppressing any cell where the population is under 20 and the count is one to three, along with any percentage that would let a reader back the number out. Survey estimates need at least 50 respondents to be released. It carries no date or statutory citation, so those fields are genuinely blank." },
  { num: 16, short: "Urban/Rural County Classification (Nov 2014)", summary: "This one-page Department of Health reference sorts all 33 New Mexico counties into four urbanization tiers: Metropolitan, Small Metro, Mixed Urban/Rural, and Rural. Bernalillo is classified Metropolitan and Catron is Rural. Dated November 2014, it is a lookup table rather than a rule and cites no statute. It is the connective tissue that explains why two counties running the identical economic template behave so differently." },
];

const SIGNALS = [
  { num: 1, signal: "Honest blanks, not invented data", detail: "Two documents genuinely carry no effective date and no statutory citation: the SPD Code of Ethics (09) and the Small Numbers Rule (15). Those cells are left blank on purpose. The model did not manufacture a plausible-looking date to fill the grid.", docs: "09, 15" },
  { num: 2, signal: "Four files were image-only scans", detail: "Documents 5, 8, 9, and 10 have no text layer at all; a plain copy-paste or keyword search returns nothing. They were run through OCR to recover their effective dates and statute cites. One quarter of the pile would have been invisible to a naive tool.", docs: "05, 08, 09, 10" },
  { num: 3, signal: "Same label, different meaning", detail: "Most rules sit under NMAC Title 1, Chapter 4 'State Procurement.' But document 3 is Title 1, Chapter 4 'Expenditure of Public Funds' under Public Finance, and document 5 is Chapter 5 entirely. The extraction kept them straight instead of collapsing every 'Chapter 4' into one bucket.", docs: "02, 03, 05, 08" },
  { num: 4, signal: "A 27-year span, all still in force", detail: "Effective dates run from April 1998 (design-build) to January 2025 (IT consolidation FAQ). Several governing rules the state relies on today are more than 25 years old. That age spread only becomes visible once the dates are pulled into one column.", docs: "05, 10, 11" },
  { num: 5, signal: "A cross-agency link that explains a number", detail: "The Health Department's urban/rural table (16) tags Bernalillo Metropolitan and Catron Rural. That predicts why the identical economic template shows Bernalillo moving 0.04% on a $7.04B base while Catron swings 34.7% on a roughly $21M base. The Health Department's Small Numbers Rule (15) is the statistical reason small populations swing. The extraction connected a health data standard to an economic report to explain an outlier.", docs: "13, 14, 15, 16" },
  { num: 6, signal: "Citations that had to be read, not scraped", detail: "Governing authorities like 'Sections 13-1-177 through 13-1-180' are buried in body prose, and OCR mangled document 10's cite into '14-1 1-7.' The extraction normalized these to clean, consistent statute references — a field that requires reading the document rather than lifting a header.", docs: "01, 08, 10" },
];

// ── Build workbook ────────────────────────────────────────────────────────────

async function build() {
  const wb = new ExcelJS.Workbook();
  wb.creator = "Brian Beals / BOB";
  wb.created = new Date();

  // ════════════════════════════════════════════════════════════════════════════
  // TAB 1 — Extracted Data
  // ════════════════════════════════════════════════════════════════════════════
  const ws1 = wb.addWorksheet("Extracted Data", {
    views: [{ showGridLines: false, zoomScale: 150, state: "frozen", ySplit: 2 }],
  });

  const headers1 = [
    "#", "Document Title", "Domain", "Document Type", "Issuing Agency",
    "Effective / Issue Date", "Governing Authority (Statute / Rule)",
    "Key Requirement", "Pages", "Text Source", "Open PDF",
  ];
  const colWidths1 = [4, 52, 12, 24, 44, 28, 54, 62, 6, 18, 10];

  colWidths1.forEach((w, i) => { ws1.getColumn(i + 1).width = w; });

  // Row 1 — Banner
  ws1.mergeCells("A1:K1");
  const banner1 = ws1.getCell("A1");
  banner1.value = "New Mexico Public Documents — Structured Extraction (16 source PDFs)";
  banner1.fill   = { type: "pattern", pattern: "solid", fgColor: { argb: NAVY } };
  banner1.font   = { name: "Arial", size: 13, bold: true, color: { argb: WHITE } };
  banner1.alignment = { horizontal: "center", vertical: "middle" };
  banner1.border = thinBorder();
  ws1.getRow(1).height = 28;

  // Row 2 — Column headers
  const hRow1 = ws1.getRow(2);
  hRow1.height = 22;
  headers1.forEach((h, i) => {
    const cell = hRow1.getCell(i + 1);
    cell.value = h;
    navyHeader(cell);
  });

  // Data rows
  EXTRACTED.forEach((doc, idx) => {
    const r = ws1.addRow([
      doc.num, doc.title, doc.domain, doc.type, doc.agency,
      doc.date, doc.statute, doc.req, doc.pages, doc.source, "Open PDF",
    ]);
    const rowIdx = idx + 1; // 1-based within data rows
    r.height = 36;

    r.eachCell({ includeEmpty: true }, (cell, colNum) => {
      cell.fill   = dataFill(rowIdx);
      cell.border = thinBorder();
      cell.font   = { name: "Arial", size: 9, bold: colNum === 1 || colNum === 2 };
      cell.alignment = { vertical: "middle", wrapText: true };
    });

    // Hyperlink on "Open PDF" column (col 11)
    const linkCell = r.getCell(11);
    const pdfPath  = `NM Procurement Docs/${doc.file}`;
    linkCell.value = { text: "Open PDF", hyperlink: pdfPath };
    linkCell.font  = { name: "Arial", size: 9, color: { argb: LINK }, underline: true };
    linkCell.alignment = { vertical: "middle", horizontal: "center" };
  });

  // ════════════════════════════════════════════════════════════════════════════
  // TAB 2 — Summaries
  // ════════════════════════════════════════════════════════════════════════════
  const ws2 = wb.addWorksheet("Summaries", {
    views: [{ showGridLines: false, zoomScale: 150, state: "frozen", ySplit: 2 }],
  });

  ws2.getColumn(1).width = 4;
  ws2.getColumn(2).width = 38;
  ws2.getColumn(3).width = 88;

  // Banner
  ws2.mergeCells("A1:C1");
  const banner2 = ws2.getCell("A1");
  banner2.value = "New Mexico Public Documents — Plain-English Summaries";
  banner2.fill   = { type: "pattern", pattern: "solid", fgColor: { argb: NAVY } };
  banner2.font   = { name: "Arial", size: 13, bold: true, color: { argb: WHITE } };
  banner2.alignment = { horizontal: "center", vertical: "middle" };
  banner2.border = thinBorder();
  ws2.getRow(1).height = 28;

  // Header row
  const hRow2 = ws2.getRow(2);
  hRow2.height = 22;
  ["#", "Document", "Plain-English Summary"].forEach((h, i) => {
    const cell = hRow2.getCell(i + 1);
    cell.value = h;
    navyHeader(cell);
  });

  // Data
  SUMMARIES.forEach((s, idx) => {
    const r = ws2.addRow([s.num, s.short, s.summary]);
    r.height = 72;
    r.eachCell({ includeEmpty: true }, (cell, colNum) => {
      cell.fill   = dataFill(idx + 1);
      cell.border = thinBorder();
      cell.font   = { name: "Arial", size: 9, bold: colNum <= 2 };
      cell.alignment = { vertical: "top", wrapText: true };
    });
  });

  // ════════════════════════════════════════════════════════════════════════════
  // TAB 3 — Cross-Document Signals
  // ════════════════════════════════════════════════════════════════════════════
  const ws3 = wb.addWorksheet("Cross-Document Signals", {
    views: [{ showGridLines: false, zoomScale: 150, state: "frozen", ySplit: 2 }],
  });

  ws3.getColumn(1).width = 4;
  ws3.getColumn(2).width = 36;
  ws3.getColumn(3).width = 88;
  ws3.getColumn(4).width = 18;

  // Banner
  ws3.mergeCells("A1:D1");
  const banner3 = ws3.getCell("A1");
  banner3.value = "New Mexico Public Documents — Cross-Document Signals";
  banner3.fill   = { type: "pattern", pattern: "solid", fgColor: { argb: NAVY } };
  banner3.font   = { name: "Arial", size: 13, bold: true, color: { argb: WHITE } };
  banner3.alignment = { horizontal: "center", vertical: "middle" };
  banner3.border = thinBorder();
  ws3.getRow(1).height = 28;

  // Header row
  const hRow3 = ws3.getRow(2);
  hRow3.height = 22;
  ["#", "Signal", "What the Extraction Caught", "Docs"].forEach((h, i) => {
    const cell = hRow3.getCell(i + 1);
    cell.value = h;
    navyHeader(cell);
  });

  // Data
  SIGNALS.forEach((sig, idx) => {
    const r = ws3.addRow([sig.num, sig.signal, sig.detail, sig.docs]);
    r.height = 80;
    r.eachCell({ includeEmpty: true }, (cell, colNum) => {
      cell.fill   = dataFill(idx + 1);
      cell.border = thinBorder();
      cell.font   = { name: "Arial", size: 9, bold: colNum <= 2 };
      cell.alignment = { vertical: "top", wrapText: true };
    });
  });

  // ── Write file ──────────────────────────────────────────────────────────────
  await wb.xlsx.writeFile(OUT);
  console.log(`✓ Written: ${OUT}`);
}

build().catch((e) => { console.error(e); process.exit(1); });
