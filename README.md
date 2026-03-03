# xlsx-md-roundtrip

A lightweight Python tool to **round-trip Excel files through a text-based representation**:

- **Export** `.xlsx` → `.md` (Markdown with an embedded YAML “blueprint”)
- **Rebuild** `.md` → `.xlsx`

The key idea is simple: **turn a binary spreadsheet into a diffable, editable, AI-friendly text artifact**, then convert it back to Excel while preserving as much layout and styling as possible.

---

## Why / What is it good for?

### 1) AI-assisted Excel editing 
Excel is hard to edit programmatically and awkward to modify via AI because it’s a binary file with lots of implicit formatting rules. By exporting to Markdown+YAML:

1. Convert your Excel file to a blueprint:
   - `.xlsx` → `.md`
2. Send the `.md` file (or a relevant excerpt) to an AI assistant:
   - Ask it to change values, formulas, formatting, merges, column widths, etc.
3. Rebuild the updated Excel:
   - `.md` → `.xlsx`

This enables workflows like:
- “Update this sheet with the new pricing table and apply the same formatting rules.”
- “Insert 10 new rows, continue formulas, keep borders and fills consistent.”
- “Rename headers, change number formats to percent, and adjust column widths.”

Because the blueprint is text, you can:
- review changes before rebuilding,
- apply changes safely and repeatedly,
- avoid hand-editing Excel for repetitive tasks.

**Tip:** for very large spreadsheets, send only the sheet(s) or ranges you want the AI to modify (copy the YAML snippet for those cells).

---

### 2) Version control and code review for spreadsheets
Storing Excel in Git is painful:
- diffs are not meaningful,
- merge conflicts are hard,
- reviews become “trust me bro, I changed something”.

With this project, you can commit the exported `.md` blueprint and get:
- readable diffs in pull requests,
- blame/traceability for when a specific cell/formula/format changed,
- easier collaboration on spreadsheet-driven artifacts.

Typical usage:
- Keep `*.md` blueprints in the repo
- Generate `*.xlsx` only for release or delivery

---

### 3) Template-driven document generation (Excel as an output format)
If your organization uses Excel as a “delivery format” (reports, price lists, checklists, test matrices), you can treat the blueprint as a template:

- Keep a baseline blueprint under version control
- Programmatically or manually adjust specific values/formulas in the YAML
- Rebuild a polished Excel file for each customer/project/release

This is especially useful when you need consistent styling and layout across many generated spreadsheets.

---

### 4) Standardizing / mass-updating formatting rules
Once content is in the blueprint, you can apply uniform transformations:
- normalize fonts and alignment across sheets,
- enforce borders for all header rows,
- remove unwanted fills,
- standardize number formats (date, currency, percent),
- align column widths.

This is helpful for “bring your spreadsheets into a standard corporate format” tasks.

---

### 5) Automation & CI workflows (Excel regeneration on demand)
Because the blueprint is deterministic input, you can rebuild Excel in automation:
- CI pipeline generates `.xlsx` from a blueprint for releases
- nightly jobs rebuild spreadsheets after data updates
- reproducible outputs (same blueprint → same spreadsheet structure)

---

### 6) Lightweight auditing & change tracking
For regulated or quality-controlled processes, you may need to answer:
- “What changed between two versions of this spreadsheet?”
- “Which formulas were modified?”
- “Did anyone change formatting that affects readability?”

Blueprint diffs provide a practical audit trail without relying on Excel’s internal history.

---

## What it preserves (current scope)

The round-trip conversion aims to preserve:

- cell values and formulas
- number formats
- fonts / fills / borders / alignment
- merged cells
- column widths / row heights
- basic sheet view properties (zoom)

It also applies a few safety rules to reduce common Excel artifacts:
- store enough color metadata (`rgb/theme/indexed/tint`) to avoid color loss and “black fill blocks”
- apply styles only to the top-left cell of merged ranges
- clear freeze/split/pane restrictions on rebuild to ensure sheets scroll normally

---

## Requirements

- Python 3.9+ (recommended)
- Dependencies:
  - `openpyxl`
  - `pyyaml`

Install:

```bash
pip install openpyxl pyyaml

---

## Usage
1) Export Excel to Markdown (YAML blueprint embedded)：
bash: python xlsx_md_roundtrip.py --xlsx input.xlsx --md blueprint.md

2) Rebuild Excel from Markdown：
bash: python xlsx_md_roundtrip.py --from-md blueprint.md --xlsx-out rebuilt.xlsx

---

## Suggested AI workflow (practical example)
1.Export:
bash: python xlsx_md_roundtrip.py --xlsx report.xlsx --md report.blueprint.md

2.Ask an AI assistant (give it the .md file) something like:
“In sheet Summary, change the header in A1 to Quarterly Report and set it bold.”
“In sheet Pricing, update the values in column C using this list, keep existing formatting.”
“Add a new row after row 10, copy formulas and borders from row 10.”

3.Rebuild:
bash: python xlsx_md_roundtrip.py --from-md report.blueprint.md --xlsx-out report.updated.xlsx

---

## Limitations / Non-goals
This tool does not aim to perfectly preserve every Excel feature (charts, pivot tables, macros/VBA, external connections, complex conditional formatting, etc.).
Very large workbooks can produce very large Markdown/YAML files; consider working per-sheet or per-range.

---

## Project Structure
xlsx_md_roundtrip.py - main script (export + rebuild)
README.md - documentation
