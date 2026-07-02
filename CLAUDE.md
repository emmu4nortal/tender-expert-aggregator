# Tender Expert Experience Aggregator

## Source of truth
`tender-expert-aggregator-spec.md` is the authoritative spec.

## Hard constraints (never violate)
- **No billed Anthropic API calls.** No `anthropic` SDK import, no API key, no HTTP calls to api.anthropic.com.
- **No local LLM.** No ollama, llama.cpp, or similar.
- Any extraction intelligence is performed by me (Claude Code) interactively under the user's subscription seat.
- No Microsoft Graph, no device-code auth, no Azure app registration.
- No Microsoft Power Platform / Power Automate / Dataverse.

## Sync root (config constant in code)
```
/Users/panu/Library/CloudStorage/OneDrive-SharedLibraries-Nortal/Public Sales - Documents
```
Note the spaces in the path. Full sync is used; do not change to selective sync.

## Excluded folders (enforced in code only — all folders ARE synced)
- `General`
- `Hinnankorotukset 2024`
- `Asiakkuussuunnitelmat`

Match on relative path from the sync root so only these top-level folders and their descendants are excluded. Keep as a config constant for easy extension.

## Tech stack
- Python 3, openpyxl, pathlib / os.walk
- No msal, no httpx, no anthropic SDK

## Current state
All milestones 0–5 complete. 872 rows in master Excel at:
`General/Referenssit/Asiantuntijat/Asiantuntijat_Master.xlsx`

Daily sync: `python run.py sync`

**Open to-do**: Milestone 6 enrichment — `technologies` and `domain_or_industry`
columns are empty for all rows. Not yet implemented.

## Data model
The unit of record is a **requirement row**: one master row = one requirement row
in the source Excel for one expert.

### JSON record fields (extraction_batch.json)
```
developer_name, role,
requirement_text,     ← verbatim source column C
evidence,             ← verbatim evidence cell (header-driven column; Formats 1/3)
technologies, domain_or_industry,
source_file_name, source_relative_path, source_sheet,
source_last_modified, extracted_date
```

### Master Excel columns (write_master.py)
Asiantuntijan nimi | Rooli tarjouksessa | Vaatimus / osaaminen |
Kokemus (alkuperäinen teksti) | Teknologiat | Toimiala |
Lähdetiedosto | Polku | Taulukko | Tiedosto muokattu | Poimittu

## Source file format types
Three formats — all use the same extraction principle (one req row → one record):
- **Format 1 (narrative)**: G = date-prefixed free text per project line. evidence = the cell under the section's "Kuvaus ... täyttymisestä" (mandatory) / "Kuvaus pisteytettäv..." (scoring) header.
- **Format 2 (parallel columns)**: Separate numbered lists in adjacent columns (clients, projects, dates, role, htp, …). evidence = each headed column stored verbatim (never split/zipped), labelled from the sheet's own **sub-header row** (the row below `Nro`; label = header's first line) so different templates' column sets/orders work and no column is dropped. Headerless columns skipped; free-text rows stored verbatim without labels.
- **Format 3 (consolidated prose)**: full prose narrative. evidence = the cell under the section's "Kuvaus ... täyttymisestä" (mandatory) / "Kuvaus pisteytettäv..." (scoring) header, verbatim.

  Formats 1/3 are **header-driven**: the evidence column is located by its header text per
  section (mandatory block, then scoring block), because the column position varies by
  template (seen in F, G, H). A sheet with no recognisable header falls back to fixed
  G (mandatory) / H (scoring). The scoring tag (`pisteytettävä`) is matched on the stem
  `pisteyt` via `_is_scoring()`.
- **Format 4 (row-per-project table)**: header-driven layout where each project is its own row; the requirement row carries the first project, continuation rows carry the rest. evidence = one labelled block per project, one `Header: value` line per **headed** column (labels taken from the source header row; headerless placeholder columns like a client's `Projekti 1` are ignored). Detected by a `Nro` header row with ≥4 headed columns in G–N (Format 1/2/3 have ≤2). Different client templates (HUS, Istekki, etc.) all use this one path.

Format detection is per-sheet in `_detect_format()`. Scans all mandatory requirement
rows before deciding — does not short-circuit on plain-text rows (e.g. "Koulutus").

Name/role is resolved per-sheet by `_find_name_role()`, trying three conventions: an explicit
`Asiantuntijan rooli:` marker (name in D); an `Asiantuntijan nimi:` label (name in the adjacent
cell, role from sheet title); and role-as-sheet-name layouts (role in col B above the `Nro`
header, name-like value in D/E). This recovers submissions that earlier produced no rows.

A Format 1/3 sheet can stack several experts, each opening with an `Asiantuntijan rooli:`
marker. When ≥2 markers are present, `_extract_sheet` splits the sheet into per-expert blocks
and runs `_extract_fmt13_rows()` on each so every requirement is attributed to the right person
(role falls back to col C when the marker cell holds only the label). Sheets with <2 markers
take the single-expert path unchanged. Stacked-roster layouts that are Format 4 (one shared
table for several listed experts) are not handled by this and remain a known gap.

## Source of truth
`extraction_batch.json` is the full, pre-dedup record set and the single source of truth. The
master Excel is a generated artifact: **`master = dedup(extraction_batch.json)`**, rebuilt in
full each time rather than patched in place. `sync` updates the batch per source file (replacing
a re-extracted file's records — 0 records removes its rows, fixing orphans) and also prunes
batch groups for files no longer present as candidates (deleted/renamed/superseded), so the
master tracks disk without a full `--all`; pruning is skipped when enumeration yields 0
candidates (sync root unavailable) to avoid wiping the batch. `write` merges a json into the
batch by source path (no deletion prune — it has no enumeration to compare against); both then
rebuild the master. Nothing reads the master back.
(Future Milestone 6 enrichment must therefore live in the batch or a content-keyed side table
joined at rebuild time — never only on the regenerated master.)

## Scripts
- `run.py` — main entry point. Commands: `sync`, `write <json>`, `status`
- `extract_requirements.py` — deterministic requirement-row extractor; `--all` processes all candidates. A workbook that fails to open propagates the error (not silently treated as 0 records), so `run.py sync` reports it failed and retries it next run instead of marking it synced; `--all` logs and skips it.
- `write_master.py` — Excel writer (11 columns, NFC path normalization). Row-level dedup key
  is `(developer_name, requirement_text, evidence)` — content only; source path/sheet are
  provenance, not identity, so the same fact across a draft and its final (or across tenders)
  collapses to one row.
- `dedupe.py` — file-level deduplication (keeps newest mtime per normalised name)
- `enumerate_candidates.py` — walks sync root, filters by keyword and excluded folders
- `config.py` — SYNC_ROOT, MASTER_PATH, STATE_FILE, EXCLUDED_TOP_LEVEL constants
- `state.json` — mtime+size per tracked file; updated by `run.py sync`
- `extraction_batch.json` — last full extraction output (1133 records → 872 unique master rows)
