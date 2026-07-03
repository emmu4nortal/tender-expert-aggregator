import os
from pathlib import Path

# Per-machine, only the home prefix differs on a Nortal Mac. Default derives from the current
# user's home so no edit is needed; override with the TENDER_SYNC_ROOT env var if your
# CloudStorage folder name differs (check ~/Library/CloudStorage/).
_DEFAULT_SYNC_ROOT = (
    Path.home() / "Library" / "CloudStorage"
    / "OneDrive-SharedLibraries-Nortal" / "Public Sales - Documents"
)
SYNC_ROOT = Path(os.environ.get("TENDER_SYNC_ROOT", _DEFAULT_SYNC_ROOT))

EXCLUDED_TOP_LEVEL = {
    "General",
    "Hinnankorotukset 2024",
    "Asiakkuussuunnitelmat",
}

CANDIDATE_KEYWORDS = [
    "cv",
    "osaaminen",
    "asiantuntij",
    "kokemus",
    "resurssi",
]

MASTER_PATH = (
    SYNC_ROOT
    / "General"
    / "Referenssit"
    / "Asiantuntijat"
    / "Asiantuntijat_Master.xlsx"
)

STATE_FILE = Path(__file__).parent / "state.json"
