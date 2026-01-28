"""
Configuration pour l'automatisation des fichiers Excel OneDrive.
"""
import os
from pathlib import Path
from dotenv import load_dotenv

# Charger les variables d'environnement
load_dotenv()

# Répertoire racine du projet
PROJECT_ROOT = Path(__file__).parent

# Répertoires
LOGS_DIR = PROJECT_ROOT / "logs"
LOGS_DIR.mkdir(exist_ok=True)

# Chemin racine OneDrive (dossier contenant les sous-dossiers de fichiers)
ONEDRIVE_BASE_PATH = Path(os.getenv("ONEDRIVE_BASE_PATH", ""))

# Configurations par fichier
SUIVI_KPIS_CONFIG = {
    "folder": "SUIVI_KPIS",
    "file_prefix": "SUIVI_KPIS",
    "date_sheet": "REPORT_HEBDO",
    "date_cell": "A1",
    "timeout_refresh": 300,
}

SUIVI_MDR_CONFIG = {
    "folder": "SUIVI_MDR",
    "file_prefix": "SUIVI_MDR",
    "date_sheet": "REPORT_MDR",
    "date_cell": "A1",
    "timeout_refresh": 300,
}

SUIVI_PMA_CONFIG = {
    "folder": "SUIVI_PMA",
    "file_prefix": "SUIVI_PMA",
    "date_sheet": "REPORT_MONDE",
    "date_cell": "A1",
    "timeout_refresh": 300,
}

SUIVI_PRODUIT_CONFIG = {
    "folder": "SUIVI_PRODUIT",
    "file_prefix": "SUIVI_PRODUIT",
    "date_sheet": "REPORT_MONDE",
    "date_cell": "A1",
    "timeout_refresh": 300,
}

# Tous les fichiers à traiter
FILE_CONFIGS = {
    "SUIVI_KPIS": SUIVI_KPIS_CONFIG,
    "SUIVI_MDR": SUIVI_MDR_CONFIG,
    "SUIVI_PMA": SUIVI_PMA_CONFIG,
    "SUIVI_PRODUIT": SUIVI_PRODUIT_CONFIG,
}
