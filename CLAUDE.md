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
All milestones 0–5 complete. 696 rows in master Excel at:
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
evidence,             ← verbatim source G cell (mandatory) or H cell (scoring)
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
- **Format 1 (narrative)**: G = date-prefixed free text per project line. evidence = G (mandatory) or H (scoring).
- **Format 2 (parallel columns)**: Separate numbered lists in G=clients, H=projects, I=dates, J=roles, K=htp. evidence = each non-empty column stored verbatim under a label (never split/zipped, so unequal column lengths can't drop or mis-pair items). Free-text rows in a Format-2 sheet are stored verbatim without labels.
- **Format 3 (consolidated prose)**: G = full prose narrative (mandatory); H = full prose (scoring). evidence = G or H verbatim.

Format detection is per-sheet in `_detect_format()`. Scans all mandatory requirement
rows before deciding — does not short-circuit on plain-text rows (e.g. "Koulutus").

## Scripts
- `run.py` — main entry point. Commands: `sync`, `write <json>`, `status`
- `extract_requirements.py` — deterministic requirement-row extractor; `--all` processes all candidates
- `write_master.py` — Excel writer (11 columns, NFC path normalization)
- `dedupe.py` — file-level deduplication (keeps newest mtime per normalised name)
- `enumerate_candidates.py` — walks sync root, filters by keyword and excluded folders
- `config.py` — SYNC_ROOT, MASTER_PATH, STATE_FILE, EXCLUDED_TOP_LEVEL constants
- `state.json` — mtime+size per tracked file; updated by `run.py sync`
- `extraction_batch.json` — last full extraction output (721 records)
