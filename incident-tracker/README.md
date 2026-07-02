# Incident report tracker (AGL fittings damaged by third parties)

Third module of the AGL tracker suite, following the same pattern as the
FTG route tracker (`run_daily.py`) and the fault report tracker
(`run_fault.py`):

| file | role |
|---|---|
| `incident_parse.py` | Parses `INCIDENT & SAFETY HAZARD REPORT` PDFs (typed text layer - no OCR needed) + QA review rules |
| `incident_tracker.py` | Updates the damaged-fitting tracker workbook; rebuilds the monthly `IR ANNEX` sheets, `QA REVIEW` and `SUMMARY` |
| `run_incident.py` | Entry point: Outlook COM fetch + `--manual` mode + `incident_processed.json` ledger (idempotent) |
| `ir_draft.py` | Drafts the monthly IR document (ADB SAFEGATE AGL SPARE PARTS REQUISITION, .docx) from the tracker |
| `config_incident.yaml` | Config - merge the `incident:` section into the suite's `config.yaml` |
| `tests/` | pytest suite against reportlab-fabricated report PDFs (fictional data) |

## What it extracts from each report

Report ID + serial (serial comes from the filename, e.g. `20251107_01_...`),
WO number, report/incident dates, reporting person and manager (matched by
header column positions), Location, Area of Incident, the three times
(occurrence / attending / rectification), every numbered fitting line
(`1.TEC102-01/015 - WO number# 43233489` - multiple fittings per report
supported), the DETAILS / OBSERVATION / ACTION / ROOT CAUSE / RESOLUTION
blocks, checked category boxes (red-marked rectangles and Wingdings glyphs,
both template variants), prepared-by / submitted-to blocks, signature
IMAGES per slot, and the annex photo count + fitting refs on the annex.

## Workbook

One row per damaged fitting in the `DAMAGED FITTINGS` sheet, upserted by
`(fitting ref, WO)` so re-runs never duplicate. If the workbook already
exists (HD's real tracker), its own header row is auto-mapped by fuzzy
matching - the team's layout is preserved and only recognised columns are
written. Rebuilt on every run:

* **`IR ANNEX <MMM-YYYY>`** - one sheet per month: the month's damaged
  fittings with incident-report references, i.e. the draft annex for the
  monthly IR (claim from ADA AGL inventory). Fill `IR NO` / `IR STATUS`
  in the data sheet once the IR is raised; the annex and summary pick
  them up on the next run.
* **`QA REVIEW`** - one row per report PDF with flags (flagged rows
  highlighted red).
* **`SUMMARY`** - per-month fittings / reports / IR raised / IR pending.

## Monthly IR draft (`ir_draft.py`)

```
python ir_draft.py NOV-2025 --ir-no 55        # fittings still without an IR NO
python ir_draft.py 2025-11 --all              # include already-claimed rows
```

Writes `ADB_IR_DRAFT_<MMM-YYYY>.docx` in the team's real requisition
layout: header (`Requisition Type : Internal Purchase`, date, `Requisition
No: ADB-IR <n>`), the items table (S/N | ERP No. | DESCRIPTION | REQUIRED
QTY | UNIT PRICE | TOTAL COST | REMARKS) with one line per fitting TYPE
and qty = fittings damaged that month, the damaged-by-unknown-vehicle
JUSTIFICATION bullets, an annex table listing every fitting with its
incident-report reference, and the REQUESTED / ACKNOWLEDGED / APPROVED
signature block.

ERP numbers, descriptions and unit prices come from the config catalog
(`incident: ir: catalog:`, keyed by fitting-type prefix: TEC, LICATC,
...). Types missing from the catalog still get a row with `<FILL>`
placeholders so the draft is always complete enough to finish by hand.
After the IR is raised, type its number into the `IR NO` column of the
data sheet - the annex/summary sheets and the next draft pick it up.

## QA flags (config `incident: review:`)

* preparer signature image missing (`require_preparer_signature`)
* ADA receiver signature missing (`require_submitted_signature`, off by
  default - ADA usually signs later)
* `Location` vs `Area of Incident` mismatch (catches copy/paste slips)
* Report ID vs report date mismatch; report filed > N days after incident
* header WO not among the fitting WOs; filename WO != document WO
* stand number outside the 100-800 series
* times missing or out of order (attended < occurred, rectified < attended)
* no annex photos; annex fitting ref not on page 1
* same WO already tracked under a different fitting ref (warning in log)

## Setup

1. Copy this folder's `.py` files into the suite folder (next to
   `common.py`, `run_daily.py`, `run_fault.py`).
2. Merge the `incident:` section of `config_incident.yaml` into
   `config.yaml`.
3. `pip install pdfplumber openpyxl pyyaml` (already present if the other
   trackers run; `pywin32` needed for the Outlook fetch).
4. Add to `run_daily.bat`:

   ```bat
   python run_incident.py >> logs\incident.log 2>&1
   ```

Manual mode (no Outlook): drop report PDFs into `incident_inbox\` and run
`python run_incident.py --manual`. Other switches: `--dry-run` (parse and
log only), `--reprocess` (ignore the ledger and re-read everything).

The module is self-contained on purpose (its own config/ledger/logging
helpers in `run_incident.py`) so it runs before any integration with
`common.py`; swapping those three helpers for the suite's shared ones is
optional and mechanical.

## Tests

```
python -m pytest tests/
```

Test PDFs are fabricated with reportlab using fictional names and WOs -
no real report is committed to this repository.
