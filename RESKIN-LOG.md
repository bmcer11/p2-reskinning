# Governance instrument reskinning — session record

Seven City of Port Phillip governance instruments moved into the corporate
instrument templates. Every reskin preserves the source text verbatim; the
changes are to presentation, plus the removal of template scaffolding that the
templates themselves tell authors to delete.

Each document was verified programmatically (every source text block present,
in order) and visually (every rendered page inspected) before delivery.

## Deliverables

| # | Instrument | Template | Result |
|---|---|---|---|
| 1 | Councillor Interaction and Request Protocol | Other | `Councillor_Interaction_and_Request_Protocol_reskinned.docx` |
| 2 | Councillor Individual Briefing Guidelines | Process | `Councillor_Individual_Briefing_Guidelines_reskinned.docx` |
| 3 | Councillor Briefing Guidelines | Process | `Councillor_Briefing_Guidelines_reskinned.docx` |
| 4 | Health and Safety Policy V2.2 | Policy | `Health_and_Safety_Policy_reskinned.docx` |
| 5 | Protective Security Policy V2.0 | Policy | `Protective_Security_Policy_reskinned.docx` |
| 6 | Body Worn Camera User Guidelines V3.0 | Process | `Body_Worn_Camera_User_Guidelines_reskinned.docx` |
| 7 | City of Port Phillip Governance Framework | Framework | `City_of_Port_Phillip_Governance_Framework_reskinned.docx` (portrait); landscape variant kept as `..._landscape.docx` |

## Method

For every instrument the template package is the base: its cover, City of Port
Phillip contact page, Document Governance and Document History tables and
Table of Contents are kept, and the instrument's content is poured in, restyled
onto the template's Heading 1/2/3 and its own automatic numbering, bullet list
and table styling.

Removed in every case, as the templates instruct: the template-instructions
table, the red guidance paragraphs, the Copilot prompt boxes and the unused
placeholder sections.

Instrument 7 was first delivered the opposite way round (source as base,
keeping its landscape design — retained as `..._landscape.docx`), then redone
on the requester's direction as the standard portrait reflow described below.

Common transformations:

- **Section structure kept from the source.** None of these instruments follow
  the templates' suggested section names, so the author's own sections and
  titles were preserved rather than forced into the template skeleton.
- **Literal heading numbers stripped** ("1. Purpose", "4.1 Practical
  Examples") so the template's automatic numbering supplies them, avoiding
  "1. 1. Purpose".
- **Tables** restyled to each template's header colour — dark grey `38322F`
  (Other), orange `F68B1F` (Process), teal `20ABAD` (Policy), magenta
  `9D2270` (Framework) — with header rows repeating across page breaks.
- **Tables of Contents** rebuilt as real Word ToC fields with correct page
  numbers, and the file set to refresh fields on open.

## Per-document decisions and defects found

**1 — Councillor Interaction and Request Protocol → Other.** Table 2 (Contact
Matrix) and its colour legend keep their original fills, which are semantic
rather than decorative. The legend was floating at a fixed page position and
was made to flow inline with Table 2. Lettered `a)`–`f)` items required
repurposing the template's unused list level 3, and removing an `isLgl` flag
that was forcing decimal display.

**2 — Councillor Individual Briefing Guidelines → Process.** The source's own
Document Governance / Review History metadata was carried over as unnumbered
headings; the "Shared understanding" callout was recoloured from a leftover
blue into the Process orange. The process flowchart and the decimal step list
were preserved.

**3 — Councillor Briefing Guidelines → Process.** The source was entirely flat:
every paragraph in one custom `NormalBullets` style, hierarchy carried only by
bold/underline. Structure was rebuilt from that formatting — bold+underlined to
Heading 1, bold-only to Heading 2, and the three timeline entries nested under
"Timelines:" to Heading 3 so the Timelines/Reporting split stayed correct.
`NormalBullets` supplies a bullet inherently, so paragraphs without a numbering
override are bulleted and plain ones opt out with `numId=0` — mapped
accordingly. **No version is stated anywhere in the source**, so the template's
`Version X.X` placeholder was deliberately left for the author.

**4 — Health and Safety Policy V2.2 → Policy.** The source carried unaccepted
tracked changes which *were* the V2.2 update (version 1 → 2.2, date →
12/08/2026, the new Psychological Health Regulations 2025 row, Strategic
Direction changed to "Trusted and High Performing Organisation"). These were
accepted first, matching the filename and cover. Two reference tables had
become adjacent siblings, which Word silently merges into one table — a spacer
paragraph was inserted between them. Old-skin decoration (cover photograph,
Purpose heading icon, hairline rule) was dropped.

**5 — Protective Security Policy V2.0 → Policy.** Already part-migrated into
this same template, so its Document Governance table was carried across with
its real values rather than reverted to blanks; those values were still styled
as grey italic placeholder text and were set to Normal as the template
instructs. Defects fixed: "2. Scope" was rendering at the wrong size (Heading 1
style overridden by direct formatting); two full sentences carried Heading 2
style by mistake and were demoted to body text so sub-numbering was correct; a
mid-document section break meant the header and footer only appeared from page
11 with page numbers restarting at 1; three conflicting inline spacing regimes
were normalised. The roles table's in-cell bullets were nearly lost to the
cell-cleanup pass and were restored.

**6 — Body Worn Camera User Guidelines V3.0 → Process.** A Guideline → Process
migration, matching the retirement of "Guideline" from the submissions app
dropdown. Both governance and history tables were fully populated and came
across whole. **Real defect fixed:** the Table of Contents was typed by hand as
plain text and was itself styled as a numbered heading, so it occupied section
1 and pushed every section one ahead. Replaced with a real ToC field; body
renumbers 1–15 with Purpose as 1. Lists used the *List Bullet* style but that
style carried no numbering definition, so they rendered as plain lines.

**7 — City of Port Phillip Governance Framework → Framework.** The source is a
designed 56-page **landscape** publication: 27 sections driving two-column
layouts, floating infographics, running section banners. First delivered with
the source as the base and the template styled over it, preserving that whole
landscape design (kept as `..._landscape.docx`); then redone as the standard
portrait reflow on the requester's direction — the delivered
`..._reskinned.docx`:

- The landscape sections and two-column layouts dissolve into a single-column
  portrait flow in reading order; the author's own styles travel with each
  paragraph, so all 161 headings keep their exact levels, unnumbered (this
  framework uses named sections, so the template's automatic clause numbering
  is stripped from the heading styles).
- Her curated Contents control is kept as a live Word ToC — same 44 entries,
  hyperlinks and bookmarks intact — with its cached page numbers re-pointed at
  the portrait pagination and its tab stop re-cut for the portrait width. The
  file still refreshes fields on open.
- The 29×10 Principles Matrix cannot survive a portrait squeeze, so its section
  sits on **landscape pages** at the table's own width and its source 10pt type
  (the template's 11pt Normal made the header words wrap mid-word), with the
  header row repeating and recoloured to the template magenta.
- The source's floating shapes were positioned for landscape spreads and either
  vanished off the portrait page (all 24 "OUR PRINCIPLES" side badges) or
  landed on top of the reflowed text (the Governance Framework Diagram, the
  principles wheel, the employee-responsibilities box). Each was re-seated: the
  badges float right of their sections with text wrapping beside them, the
  five-shape diagram is restacked — house centred, the two arrow-callouts
  converging beneath it — and the wheel and box sit centred in the flow,
  reserving their own space.
- 72 body paragraphs carried direct right indents of 3,000–5,200 twips — the
  author's way of fitting text into the narrow landscape columns. They pinch
  portrait paragraphs to half width for no reason, so indents above 700 twips
  were dropped.
- The teal family (`2EBBB8`, `009999`, `00A3AD`, the shapes' `5DA5AF` and
  others) maps to the template magenta in both ordinary text and DrawingML
  shape fills; heading runs drop their direct colours so the template styles
  govern. The VML duplicates of reworked shapes were removed so a repositioned
  shape cannot disagree with a stale fallback.

## Known outstanding items

- **3 — Councillor Briefing Guidelines:** version number still the template
  placeholder; the source states none.
- **4 — Health and Safety Policy:** the source's section 4 "Internal Overview"
  governance metadata sits in the body while the template's Document Governance
  and History tables keep their placeholders. The field lists are not a clean
  1:1, so merging them needs a mapping decision.
- **5 — Protective Security Policy:** "Date last full refresh approval" and
  "Next full refresh date" remain "Select date"; the Document History table is
  still template placeholders. In the Building Safety & Security Committee row,
  one sentence is split across bullets mid-sentence — faithful to the source,
  but likely an editorial error worth fixing.
- **7 — Governance Framework (portrait):** no version is stated anywhere in
  the source, so the cover keeps the template's `Version X.X` placeholder; the
  Document Governance and Document History tables are blank placeholders — the
  source had none. The governance wheel and the principle icons are raster
  artwork, so their multicoloured palette cannot be recoloured to the
  template's magenta. Where a landscape spread paired text and imagery
  side-by-side, the portrait flow reads left column then right column. In the
  landscape variant (`..._landscape.docx`) the earlier caveats stand: banner
  teal baked into artwork, source footers and cover.

## Related change outside the documents

The Governance Instrument Submissions Power App had a document-type dropdown
still offering "Guideline" and "Procedure" after the list was updated to use
"Process". Cause: two dropdowns exist, and the edit had been made to the one on
`UpdateCoreDetails`. The control rendering on the Create screen is
`DropdownDocType_UpdateCore_1` — copy-pasted from the Update screen and never
renamed, so it reads like an Update control while sitting in Create. Fix is to
update its `Items` on `CreateCoreDetails`, and to add "Process" to the
underlying choice column before publishing, since the submit flow writes the
value as a string into a column read back as `'Document type'.Value`.

## Scripts

`scripts/` holds the build script for each document, named for its
instrument and target template. Each is self-contained: it unpacks the
template and source `.docx`, performs the transformation, and writes the
result. `_outline.py` dumps a readable structure of any `word/document.xml`
(styles, numbering levels, table shapes) and was the main analysis tool;
`_verify_framework.py` is the fidelity checker used on the landscape
framework build.

The portrait framework redo is `07b_governance_framework_portrait.py`, a
two-pass build: build, render, harvest the ToC page numbers that LibreOffice
resolves from the PAGEREF bookmarks (`07b_harvest_toc_pages.py` writes
`toc_pages8.json`), then build again so the cached numbers match the portrait
pagination. `07b_verify_portrait.py` is its fidelity checker: every source
body block present in order, ToC entries and heading levels identical.

The scripts expect `template_unpacked/` and `instrument_unpacked/` beside them,
so they are a record of exactly what was done rather than a turnkey pipeline.
