# Specification: Tender Expert Experience Aggregator (v2)

> **Status**: v2 — all milestones 0–5 complete. 696 rows in master Excel, 549 source
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
- `requirement_text` (column C) and `requirement_number` (column B) were never captured.
- `short_description` was constructed, not verbatim.

### JSON record fields (`extraction_batch.json`)

```
developer_name        — expert's full name
role                  — role declared in the tender file
requirement_text      — verbatim column C text (never transformed)
evidence              — verbatim source cell content (never transformed; see format rules)
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

`requirement_text` and `evidence` are verbatim source cell values — never constructed or transformed.

Row height: 60 pt (evidence cells are multi-line). Freeze pane at B2. Auto-filter on row 1.

---

## 7. Source file format types

Source files use one of three formats. `extract_requirements.py` detects the format
per sheet and handles each accordingly.

### Format 1 — Narrative

- `[B]` = requirement number
- `[C]` = requirement text
- `[D]` = "Pakollinen" or scoring scale text
- `[G]` = full narrative evidence (date-prefixed lines, one project per line)
- `[H]` = scoring evidence (for scoring rows)

`evidence` = `[G]` for mandatory rows, `[H]` for scoring rows.

Detection: `[G]` cell starts with a date pattern `\d{1,2}/\d{4}`.

### Format 2 — Parallel columns

Each requirement row has separate numbered lists:
- `[G]` = numbered client list
- `[H]` = numbered project name list
- `[I]` = numbered date list
- `[J]` = numbered role list
- `[K]` = numbered HTP list
- `[L]` = description (often empty)

`evidence` stores each non-empty source column **verbatim** under a label, never split
or zipped:
```
Asiakkaat:
1. SSAB Europe Oy
2. ...

Projektit:
1. ...

Ajankohta:
...
```
Zipping into per-item lines was abandoned: the columns frequently have unequal lengths
(a missing client/role, a wrapped cell, or a restarted sub-block), and positional zipping
silently dropped fields or could mis-pair the wrong client with the wrong project. Storing
the columns verbatim preserves the source exactly and is faithful to the "verbatim" rule.

A requirement row inside a Format-2 sheet whose primary column is **not** a numbered list
(e.g. a free-text `Koulutus` or certification answer) is stored verbatim without labels,
scanning the content columns for the answer.

Scoring rows are shifted one column right (`[G]` = numeric score, clients start at `[H]`).

Detection: both `[G]` and `[H]` cells start with `\d+\.\s*\S` (numbered list pattern).

### Format 3 — Consolidated prose

- Mandatory rows: `[G]` = full prose narrative per client+project+dates+role+htp+description
- Scoring rows: `[G]` = numeric score, `[H]` = full prose narrative

`evidence` = `[G]` for mandatory rows, `[H]` for scoring rows.

Detection: fallback — neither Format 1 nor Format 2 pattern found in any mandatory row.

---

## 8. Deduplication

### File-level (before extraction)
See section 5. Run by `dedupe.py` via `enumerate_candidates()`.

### Row-level (after extraction, before write)
Dedup key: `(developer_name, requirement_text, source_relative_path, source_sheet)`

Keep the record with the lexicographically largest `source_last_modified` when duplicates
exist. This handles:
- Same expert extracted twice from the same file (idempotent re-runs)
- A file re-extracted after an edit (stale rows are first removed by path, then new rows
  are deduped)

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
| `short_description` was constructed, not verbatim | `evidence` = verbatim source cell, never constructed |
| `requirement_text` never captured | Column C captured as `requirement_text` |
| `extraction_batches/` batch files have no traceability | New flow: one `extraction_batch.json` per sync run; source path in each record |
| Dedup key used constructed fields | New dedup key uses structural fields (path, sheet, req number) |
| Full re-extraction required on schema change | After migration, incremental sync runs only changed files |
| content_batch.json overwritten on each scan | Eliminated: new flow writes directly to extraction_batch.json |

---

## 11. Milestones (remaining work)

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
- Update `dedupe()` key to `(developer_name, requirement_number, source_relative_path, source_sheet)`.
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
  1. Find developer_name: scan col D rows 10–20 for first non-empty, non-header value.
     Candidate header markers: "Asiantuntijan N rooli:" in col B same row.
  2. Find role: col B of that same row, strip "Asiantuntijan N rooli: " prefix.
  3. Detect format (Format 1 / 2 / 3) by inspecting the first non-empty [G] cell
     in a requirement row (see section 7 detection rules).
  4. Walk rows:
     a. If col B matches r'^[A-EZ]?\s?\d+' → this is a requirement row.
     b. requirement_number = col B value (stripped)
     c. requirement_text   = col C value
     d. requirement_type   = "pakollinen" if col D contains "pakollinen" (case-insensitive),
                             else "pisteytettävä"
     e. Build evidence per format (section 7).
     f. points = int(col G) if scoring row and col G is numeric, else null.
     g. Skip row if evidence is None or is a template instruction (starts with
        "Kuvaus", "Asiakkaat", "Toimeksiantojen").
     h. Skip row if requirement_number is empty/None.
  5. technologies and domain_or_industry: leave as "" — filled in enrichment step.
  6. Emit record with all fields from section 6.
```

CLI:
```
python extract_requirements.py <path1> [<path2> ...]   # extract specific files
python extract_requirements.py --all                   # extract all candidates in sync root
```

Output: `extraction_batch.json` (creates or overwrites).

**Exit criteria**: Script runs on 3 test files (one per format) without error; output
records have non-empty `requirement_number`, `requirement_text`, and `evidence` for all rows.

---

### Milestone 3 — Sample extraction and user verification

**Goal**: Extract 3–5 representative files (at least one per format type), write a sample
master Excel, and let the user verify the output before the full run.

Steps:
1. Run `extract_requirements.py` on sample files covering all three format types.
2. Run `python run.py write extraction_batch.json` → writes sample master.
3. User opens `Asiantuntijat_Master.xlsx` in OneDrive and inspects:
   - Are `requirement_text` values correct (verbatim column C)?
   - Is `evidence` verbatim (not constructed)?
   - Are `requirement_type` values correctly classified?
   - Are row heights and column widths readable?
4. User confirms or requests fixes before full run.

**Exit criteria**: User explicitly approves the sample output. No proceeding to Milestone 4
without approval.

---

### Milestone 4 — Full extraction (all source files)

**Goal**: Re-extract all source files found in the sync root and rebuild the master
from scratch using the new schema.

Steps:
1. Run `python run.py sync --all` — enumerates all candidates, extracts every file via
   `extract_requirements.py`, and writes the master.
   - Cross-check: source paths in `extraction_batch_v1_backup.json` confirm which files
     were in the v1 master; all should still be present in the sync root.
2. Verify: `python -c "import json; recs=json.load(open('extraction_batch.json')); print(len(recs), 'records,', sum(1 for r in recs if not r.get('requirement_number')), 'missing req numbers')"` → second number must be 0.
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
- Add `cmd_sync(force_all=False)` implementing the algorithm in section 9.
- Wire `sync` and `sync --all` in `main()`.
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
- `python run.py sync --all` re-extracts all files and produces same row count.

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
