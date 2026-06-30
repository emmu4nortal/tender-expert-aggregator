# Specification: Tender Expert Experience Aggregator (v2)

> **Status**: v2 — milestones 0–5 complete. 839 rows in master Excel, 549 source
> files tracked. Milestone 6 (enrichment) is optional and not yet implemented.

---

## Quick start

### Setup & how to run
1. OneDrive is running and the shared library is fully synced (check the OneDrive menu bar icon — it should show no pending activity)
2. Open Terminal (macOS: Cmd+Space → type "Terminal" → Enter)
3. Copy each step to terminal and run

### Daily sync — pick up new or changed files and update the master Excel
1. cd "/Users/panu/Documents/Claude Code Projects/Tender Expert Experience Aggregator"
2. python3 run.py sync

The master Excel at `General/Referenssit/Asiantuntijat/Asiantuntijat_Master.xlsx` is updated automatically and syncs back to OneDrive.

### Check status — show tracked file count, last run time, and master row count
1. cd "/Users/panu/Documents/Claude Code Projects/Tender Expert Experience Aggregator"
2. python3 run.py status

### Full re-extraction — rebuild master from scratch (use after a code change)
1. cd "/Users/panu/Documents/Claude Code Projects/Tender Expert Experience Aggregator"
2. python3 extract_requirements.py --all
3. python3 run.py write extraction_batch.json

---

## 1. Purpose

An on-demand tool that scans a locally synced copy of Nortal's public-tender SharePoint
library, finds the Excel files listing named experts submitted in past tenders, extracts
that experience into structured rows keyed to the tender's requirement rows, and maintains
a single master Excel. The goal is to let a Delivery Director search one workbook to
identify candidate experts for new public tenders.

Personal productivity tool run by one user on macOS. Simple and maintainable over clever.

---

## 2. Hard constraints (unchanged from v1)

- Runs on macOS, on demand (`python run.py sync`).
- **No metered/billed LLM calls.** No `anthropic` SDK, no API key, no `api.anthropic.com`
  calls. No local LLM (no ollama, llama.cpp, or similar). Any intelligence needed during
  extraction runs interactively inside Claude Code on the user's subscription seat.
- No Microsoft Power Platform, Power Automate, or Dataverse.
- No Microsoft Graph, no device-code auth, no Azure app registration.
- Do not hardcode secrets (none are needed).

---

## 3. Source: locally synced SharePoint library (unchanged)

Sync root (config constant):
```
/Users/panu/Library/CloudStorage/OneDrive-SharedLibraries-Nortal/Public Sales - Documents
```

Full sync is used; all folders are synced. New top-level folders added to SharePoint later
appear automatically under the sync root. The excluded folders are skipped in code only.

Excluded top-level folders (config constant, skip during walk):
- `General`
- `Hinnankorotukset 2024`
- `Asiakkuussuunnitelmat`

Master output path (also excluded from scanning):
```
General/Referenssit/Asiantuntijat/Asiantuntijat_Master.xlsx
```

---

## 4. Candidate filter (unchanged)

An item qualifies only if ALL hold:

- It is a file, not a directory.
- Its name ends with `.xlsx`, case-insensitive.
- Its name contains (case-insensitive) at least one of: `cv`, `osaaminen`, `asiantuntij`,
  `kokemus`, `resurssi`.
- It is NOT the master output file.

---

## 5. File-level dedupe (unchanged)

Group candidate files in the same parent directory with the same normalised name; keep
only the most recently modified. Normalise by stripping version/noise tokens:
`VANHA`, `LUONNOS`, `draft`, `KORJ`, `KORJATTU`, copy markers `(1)`, trailing dates,
version suffixes `_ver1`, `V2`. Do not collapse files whose subject words differ.

---

## 6. Data model (v2 — requirement-row-centric)

**Unit of record**: one master row = one requirement row in the source Excel for one expert.
Do not split or aggregate project history that appears as one block in the source.

This replaces the old project-centric model where one record = one project experience.
The old model was wrong because:
- Source files are structured by requirement, not by project.
- Splitting one requirement's multi-project evidence into multiple records lost context.
- `requirement_text` (column C) was never captured.
- `short_description` was constructed, not verbatim.

### JSON record fields (`extraction_batch.json`)

```
developer_name        — expert's full name
role                  — role declared in the tender file
requirement_text      — verbatim column C text (never transformed)
evidence              — source experience text; verbatim cell for Formats 1/3, verbatim column values under labels for Formats 2/4 (see §7)
technologies          — comma-separated tech tags (enrichment step; may be empty)
domain_or_industry    — industry tags (enrichment step; may be empty)
source_file_name      — basename of source file
source_relative_path  — path relative to sync root (stable dedup/update key)
source_sheet          — sheet name within the workbook
source_last_modified  — file mtime at extraction time (YYYY-MM-DD HH:MM)
extracted_date        — date extraction was run (YYYY-MM-DD)
```

### Master Excel columns

| # | Key | Finnish label | Width |
|---|-----|---------------|-------|
| A | developer_name | Asiantuntijan nimi | 22 |
| B | role | Rooli tarjouksessa | 30 |
| C | requirement_text | Vaatimus / osaaminen | 60 |
| D | evidence | Kokemus (alkuperäinen teksti) | 100 |
| E | technologies | Teknologiat | 40 |
| F | domain_or_industry | Toimiala | 28 |
| G | source_file_name | Lähdetiedosto | 40 |
| H | source_relative_path | Polku | 55 |
| I | source_sheet | Taulukko | 20 |
| J | source_last_modified | Tiedosto muokattu | 20 |
| K | extracted_date | Poimittu | 14 |

`requirement_text` is always the verbatim column-C value. `evidence` is the verbatim source
cell for Formats 1/3; for the column/table Formats 2/4 it is the source column values stored
verbatim under labels (never zipped or paraphrased) — see §7.

Row height: 60 pt (evidence cells are multi-line). Freeze pane at B2. Auto-filter on row 1.

---

## 7. Source file format types

Source files use one of four formats. `extract_requirements.py` detects the format
per sheet and handles each accordingly.

### Format 1 — Narrative

- `[B]` = requirement number
- `[C]` = requirement text
- `[D]` = "Pakollinen" or scoring scale text
- evidence column = the one headed `Kuvaus ... täyttymisestä` (mandatory) / `Kuvaus
  pisteytettäv...` (scoring) — narrative evidence, date-prefixed lines, one project per line

`evidence` = the cell under the section's evidence header (see "Header-driven columns" below).

Detection: evidence cell starts with a date pattern `\d{1,2}/\d{4}`.

### Format 2 — Parallel columns

Each requirement row has separate numbered lists in adjacent columns (client list, project
list, dates, role, htp, …). The set and order of columns varies by client template, so the
columns are **labelled from the sheet's own sub-header row** rather than by fixed position.

Each requirement section has a `Nro` header row followed by a **sub-header row** whose
headed columns name the lists (the `Nro` row's `[G]` is just a section title like
`Kuvaus vähimmäisvaatimuksen täyttymisestä`). The label is the header's first line.

`evidence` stores each headed column **verbatim** under its header label, never split or
zipped (zipping was abandoned because unequal column lengths silently dropped or mis-paired
items). Headerless columns are skipped. Example:
```
Asiakkaat:
1. SSAB Europe Oy
2. ...

Toimeksiantojen nimet:
1. ...

Toimeksiantojen ajankohdat:
...
```
Driving labels from the sub-header row (instead of hardcoded positions) means a template
with extra columns — e.g. a `Toimeksiantojen tarkennukset` description column and a
`Projektin kokonaislaajuus` scope column — is labelled correctly and no column is dropped.

A requirement row whose first headed column is **not** a numbered list (e.g. a free-text
`Koulutus` or certification answer) is stored verbatim without labels, joining the headed
columns' values.

Scoring sections shift one column right (`[G]` = numeric score), which the sub-header row
reflects automatically — its headed columns simply start at `[H]`.

Detection: both `[G]` and `[H]` cells of a requirement row start with `\d+\.\s*\S`
(numbered list pattern).

### Format 3 — Consolidated prose

- Mandatory rows: evidence cell = full prose narrative per client+project+dates+role+htp+description
- Scoring rows: a numeric score column, plus the evidence cell = full prose narrative

`evidence` = the cell under the section's evidence header (see "Header-driven columns" below).

Detection: fallback — neither Format 1 nor Format 2 pattern found in any mandatory row.

### Header-driven columns (Formats 1 and 3)

The evidence column is **not** fixed at `[G]`/`[H]`: across templates the mandatory-evidence
header `Kuvaus vähimmäisvaatimuksen täyttymisestä` is seen in `[F]`, `[G]` or `[H]`, and the
requirement/tag columns shift too. So Formats 1/3 locate the evidence column by its header
text per section — a sheet has a mandatory block then (usually) a scoring block, each opened
by its own header row — and read evidence from that column for the rows in that block. A sheet
with no recognisable evidence header falls back to fixed `[G]` (mandatory) / `[H]` (scoring),
preserving older files. The scoring tag is matched on the stem `pisteyt` (`_is_scoring()`).

A row whose evidence merely echoes its requirement text is dropped (a non-answer, seen in
listing/comparison sheets that have no real evidence column).

### Format 4 — Row-per-project table

A header-driven layout where the sheet has a labelled column-header row and **each past
project occupies its own physical row** (not a numbered list inside one cell). Different
tendering clients use different column sets (e.g. client / contact / dates / scope / workload
/ role, or project description / scope / dates split across columns), so extraction is driven
by the actual headers rather than fixed positions.

- The header row has `[B]` = `Nro` and several headed columns in `[G]`–`[N]`.
- A requirement row (`[B]` = number, `[C]` = text) carries the first project; additional
  projects follow in rows with empty/non-number `[B]`. Extraction gathers the requirement
  row plus those continuation rows (stopping at the next requirement row, repeated header,
  or a row with no value in any headed column).

`evidence` = one labelled block per project, one `Header: value` line per **headed** column,
values verbatim, empty cells omitted. Columns with **no header** are ignored — this is how a
client's placeholder column (e.g. an unlabelled `[G]` holding `Projekti 1`, `Asiakkuus 1`) is
dropped while the real data in the headed columns is kept. Example:
```
Toimeksiantaja: Stora Enso Oyj
Ajankohta kk/vv-kk/vv: 11/2024 - Jatkuu edelleen
Projektin laajuus (htp): 350
...

Toimeksiantaja: ...
```

Detection: a `Nro` header row with at least 4 headed columns in `[G]`–`[N]`. Genuine
Format 1/2/3 sheets have at most 1–2 headed columns there, so the threshold separates them
cleanly and works for any client's column naming.

---

## 8. Deduplication

### File-level (before extraction)
See section 5. Run by `dedupe.py` via `enumerate_candidates()`.

### Row-level (after extraction, before write)
Dedup key: `(developer_name, requirement_text, evidence)` — content only.

The master collects **unique** expert experience, so the same fact appearing in several
source files (a draft and its final submission, or one expert proposed across tenders) is one
row, not many. Source path/sheet are provenance, **not** identity, and are deliberately *not*
in the key.

`evidence` is part of the key because one requirement can appear in a sheet as both a
mandatory (`Pakollinen`) and a scoring (`Pisteytettävä`) row — same `requirement_text`,
different evidence. They must stay as two rows, and `evidence` is what distinguishes them.
`developer_name` already separates different experts.

Keep the record with the lexicographically largest `source_last_modified` when duplicates
exist. This handles:
- Same expert extracted twice from the same file (idempotent re-runs)
- A file re-extracted after an edit (stale rows are first removed by path, then new rows
  are deduped)
- The same content in a draft and a final (or across tenders) — collapsed to one row

---

## 9. Incremental sync (daily on-demand)

The tool must support daily on-demand runs that only process new and changed files.
Full re-extraction of all files on every run is not acceptable.

### State file (`state.json`)

```json
{
  "files": {
    "Customer/Tender/Jättö/Liite 2 Asiantuntijat.xlsx": {
      "mtime": 1782389335.144,
      "size": 50218
    }
  },
  "last_run": "2026-06-26T10:00:00"
}
```

### `python run.py sync` — the daily command

1. **Enumerate** candidates via `enumerate_candidates()` + file-level dedupe (`dedupe.py`).
2. **Classify** each candidate: new (not in state), changed (mtime differs), unchanged.
3. **Extract** only new and changed files using `extract_file()` from `extract_requirements.py`.
   Records are held in memory — not written to `extraction_batch.json`.
4. **Write** master via `_merge_and_write()`:
   - Remove stale rows from master for the re-processed source paths.
   - Merge new records.
   - Dedup on key.
   - Write Excel.
5. **Update** `state.json`: record new mtime/size for successfully processed files; update `last_run`.
   Failed files are reported but not marked as synced — they will be retried next run.

If no files have changed since last run: print a summary and exit without writing.

### `python run.py status`

Print: tracked file count, last run timestamp, master row count, count of new/changed
files pending extraction.

---

## 10. Lessons learned from v1 build (addressed in v2)

| v1 Problem | v2 Fix |
|---|---|
| Flat cell dump lost row-level context | `extract_requirements.py` reads Excel row by row; groups B+C+D+G/H per requirement |
| Multiple manual Claude Code sessions required for initial extraction | Deterministic script; no agent sessions for extraction |
| `short_description` was constructed, not verbatim | `evidence` = source cell verbatim (Formats 1/3) or verbatim column values under labels (Formats 2/4) — never paraphrased |
| `requirement_text` never captured | Column C captured as `requirement_text` |
| `extraction_batches/` batch files have no traceability | New flow: one `extraction_batch.json` per sync run; source path in each record |
| Dedup key used constructed fields | New dedup key is content only (developer, requirement_text, evidence) |
| Full re-extraction required on schema change | After migration, incremental sync runs only changed files |
| content_batch.json overwritten on each scan | Eliminated: new flow writes directly to extraction_batch.json |

---

## 11. Milestones

Milestones 0–5 are complete and are kept below as a build record. **Milestone 6
(enrichment) is the only open/remaining milestone.**

### Milestone 0 — Codebase cleanup

**Goal**: Remove all v1 artefacts so the working tree is clean before new code is written.

Files to DELETE:
- `update_descriptions.py` — backfill script for old schema, no longer needed
- `sample_content.json` — output of old extract_sample.py sample run
- `sample_extraction.json` — old schema sample
- `extraction_batches/` — entire directory (10 batch files + 10 compact batches + compacted_all.json)
- `content_batch.json` — flat cell dump from old scan flow, no longer the intermediate format

Files to RENAME (backup, do not delete):
- `extraction_batch.json` → `extraction_batch_v1_backup.json`
- `Asiantuntijat_Master.xlsx` in OneDrive → `Asiantuntijat_Master_v1_backup.xlsx`
  (rename in OneDrive so it stays out of the master path)

Files to KEEP and UPDATE:
- `config.py` — no changes needed
- `enumerate_candidates.py` — no changes needed
- `dedupe.py` — no changes needed
- `write_master.py` — schema update (Milestone 1)
- `run.py` — add `sync` command (Milestone 5)
- `state.json` — keep as-is; records existing file mtimes (still valid for incremental runs)

**Exit criteria**: `ls` of project directory shows no `extraction_batches/`, no
`sample_content.json`, no `update_descriptions.py`; master backup exists in OneDrive.

---

### Milestone 1 — Update `write_master.py` (new schema)

**Goal**: The Excel writer reflects the new requirement-row schema.

Changes:
- Replace `COLUMNS`, `COLUMN_WIDTHS`, `HEADER_LABELS` with the table in section 6.
- Row height for data rows: 60 pt.
- Update `dedupe()` key to `(developer_name, requirement_text, evidence)` — content only, so the same fact across files/tenders collapses to one row. (`requirement_number` is not stored — see Milestone 2 — so the key uses `evidence` to separate a requirement's mandatory and scoring rows.)
- `load_records()`: no logic changes needed; `rec.get(col_key)` already handles new fields.

**Exit criteria**: `write_master.py` imports cleanly; column constants match section 6 exactly.

---

### Milestone 1.5 — Layout confirmation (USER REVIEW GATE)

**Goal**: Confirm the master Excel column layout before writing any data.

Claude Code presents the proposed layout (column letters, Finnish labels, widths) as a
table and asks the user to confirm or request changes. **Do not proceed to Milestone 2
until the user has approved the layout.**

This is the last easy point to change column names or order before records are generated.

---

### Milestone 2 — Write `extract_requirements.py`

**Goal**: A deterministic script that reads source Excel files directly and emits
requirement-row records in the new schema.

Algorithm per file:

```
For each expert sheet (skip helper sheets: "pisteet", "data-", "ohjeet"):
  1. Find developer_name + role via `_find_name_role()` — three conventions: an
     "Asiantuntijan N rooli:" marker (name in col D), an "Asiantuntijan nimi:" label (name in
     the adjacent cell, role from sheet title), or role-as-sheet-name (role in col B above the
     `Nro` header, name-like value in D/E). See section 7.
  2. Multi-expert sheets (Format 1/3): if ≥2 `Asiantuntijan rooli:` markers are present, split
     the sheet into per-expert blocks (marker_i .. marker_{i+1}) and extract each block with its
     own name/role, so every requirement is attributed to the right person. <2 markers → one
     expert over the whole sheet. (A Format-4 roster sheet listing several experts over one
     shared table is a known gap, not handled here.)
  3. Detect format (Format 1 / 2 / 3 / 4) by inspecting requirement rows / the column
     header (see section 7 detection rules).
  4. Walk rows:
     a. If col B matches r'^[A-EZ]?\s?\d+' → this is a requirement row (this gate is the
        requirement-number check; the number itself is not stored).
     b. requirement_text = col C value.
     c. Formats 1/3: evidence is read from the section's header-driven evidence column
        (`Kuvaus ... täyttymisestä` / `Kuvaus pisteytettäv...`), falling back to G/H; the
        scoring tag (`pisteyt` stem) is used only when no header is found. See section 7.
     d. Build evidence per format (section 7).
     e. Skip row if evidence is None or is a template instruction (starts with
        "Kuvaus", "Asiakkaat", "Toimeksiantojen").
  5. technologies and domain_or_industry: leave as "" — filled in enrichment step.
  6. Emit record with all fields from section 6. (requirement_number, requirement_type and
     points are NOT stored — they are bookkeeping only and would not help the reader.)
```

CLI:
```
python extract_requirements.py <path1> [<path2> ...]   # extract specific files
python extract_requirements.py --all                   # extract all candidates in sync root
```

Output: `extraction_batch.json` (creates or overwrites).

A workbook that fails to open (locked, corrupt) propagates the error rather than returning an
empty record list, so it is distinguishable from a genuinely empty sheet: `run.py sync` reports
it failed and does not mark it synced (retried next run); `--all` logs and skips it.

**Exit criteria**: Script runs on test files (one per format) without error; output
records have non-empty `requirement_text` and `evidence` for all rows.

---

### Milestone 3 — Sample extraction and user verification

**Goal**: Extract 3–5 representative files (at least one per format type), write a sample
master Excel, and let the user verify the output before the full run.

Steps:
1. Run `extract_requirements.py` on sample files covering all format types.
2. Run `python run.py write extraction_batch.json` → writes sample master.
3. User opens `Asiantuntijat_Master.xlsx` in OneDrive and inspects:
   - Are `requirement_text` values correct (verbatim column C)?
   - Is `evidence` correct (verbatim cell for narrative/prose; correctly labelled columns for the table formats)?
   - Are row heights and column widths readable?
4. User confirms or requests fixes before full run.

**Exit criteria**: User explicitly approves the sample output. No proceeding to Milestone 4
without approval.

---

### Milestone 4 — Full extraction (all source files)

**Goal**: Re-extract all source files found in the sync root and rebuild the master
from scratch using the new schema.

Steps:
1. Run `python extract_requirements.py --all` then `python run.py write extraction_batch.json`
   — extracts every candidate and rebuilds the master from scratch.
   - Cross-check: source paths in `extraction_batch_v1_backup.json` confirm which files
     were in the v1 master; all should still be present in the sync root.
2. Verify: `python -c "import json; recs=json.load(open('extraction_batch.json')); print(len(recs), 'records,', sum(1 for r in recs if not r.get('evidence')), 'missing evidence')"` → second number must be 0.
3. Spot-checks:
   - Essi Suo-Heikki: should have ~5 rows (D1, D3, D8, D9, D10).
   - Mikko Keiho / Valtiokonttori: `evidence` must contain full date-prefixed narrative.
   - No empty `evidence` fields.

**Exit criteria**: Master written, spot-checks pass, user confirms row count is plausible.

---

### Milestone 5 — Incremental sync (`run.py sync`)

**Goal**: A single command for daily on-demand updates that only processes new and changed
files. No full re-extraction, no manual steps.

Changes to `run.py`:
- Add `cmd_sync()` implementing the algorithm in section 9.
- Wire `sync` in `main()`. (Full re-extraction is `extract_requirements.py --all` +
  `run.py write`, not a `sync` flag — there is no `sync --all`.)
- Remove or deprecate `scan` command (it produces `content_batch.json` for the old agent
  flow which no longer exists).
- Update `status` to show pending new/changed file count.

The sync command must be fully unattended (no Claude Code session needed) as long as
`extract_requirements.py` handles the files involved. If a new source file uses an
unexpected format, `extract_requirements.py` should emit a warning and skip the file
rather than crashing, and `cmd_sync()` should report skipped files to the user.

**Exit criteria**:
- `python run.py sync` on a clean run → "0 files changed, nothing to do."
- Manually touch a source file (update its mtime) → sync detects and re-extracts it.
- `python extract_requirements.py --all` + `python run.py write extraction_batch.json`
  rebuilds all files and produces the same row count.

---

### Milestone 6 — Optional enrichment pass (technologies / domain_or_industry)

**Goal**: Populate the `technologies` and `domain_or_industry` fields, which require
reading the `evidence` text and tagging it.

This step is explicitly Claude Code interactive. Since it requires intelligence
(no deterministic mapping is possible) and cannot use billed API:
- Claude Code reads `extraction_batch.json` in batches.
- For each record, reads `evidence` and suggests comma-separated technology tags
  and domain/industry tags.
- Writes enriched records back to `extraction_batch.json`.
- Re-runs `python run.py write extraction_batch.json` to update master.

This milestone is **optional** for the daily sync. The master is fully usable without
it (evidence cells contain all information; technologies/domain are convenience filters).

**Exit criteria**: At least the 24 source files' records have non-empty `technologies`
and `domain_or_industry` values.

---

## 12. Non-goals (unchanged)

Billed or local LLMs, the Graph API and authentication, Power Platform/Dataverse,
real-time triggers, hosting or scheduling, and non-xlsx file types are explicitly out
of scope.
