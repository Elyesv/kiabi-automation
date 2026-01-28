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

# Configuration SUIVI_KPIS
SUIVI_KPIS_CONFIG = {
    "folder": "SUIVI_KPIS",
    "file_prefix": "SUIVI_KPIS",
    "date_sheet": "REPORT_HEBDO",
    "date_cell": "A1",
    "verification_sheet": "SUIVI_JOUR",
    "timeout_refresh": 300,
}

# Mapping des configurations par type de fichier (pour extensions futures)
FILE_CONFIGS = {
    "SUIVI_KPIS": SUIVI_KPIS_CONFIG,
}
