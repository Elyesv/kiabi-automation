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
    "file_ext": ".xlsm",
    "date_sheet": "REPORT_MONDE",
    "date_cell": "A1",
    "timeout_refresh": 300,
}

SUIVI_PRODUIT_CONFIG = {
    "folder": "SUIVI_PRODUIT",
    "file_prefix": "SUIVI_PRODUIT",
    "file_ext": ".xlsm",
    "date_sheet": "REPORT_MONDE",
    "date_cell": "A1",
    "timeout_refresh": 300,
}

SUIVI_CRM_CONFIG = {
    "folder": "SUIVI_CRM",
    "file_prefix": "SUIVI_CRM",
    "timeout_refresh": 300,
    # Requêtes Power Query à mettre à jour
    "queries": {
        # Requêtes selligent : mettre à jour le numéro de semaine dans le chemin
        "selligent_all": {"type": "selligent"},
        "selligent_all_histo": {"type": "selligent"},
        # Requêtes piano : mettre à jour les dates start/end (+7 jours)
        "piano_all": {"type": "piano"},
        "piano_all_histo": {"type": "piano"},
    },
}

# Fichiers avec mise à jour date + refresh (lancés par update_all.py)
FILE_CONFIGS = {
    "SUIVI_KPIS": SUIVI_KPIS_CONFIG,
    "SUIVI_MDR": SUIVI_MDR_CONFIG,
    "SUIVI_PMA": SUIVI_PMA_CONFIG,
    "SUIVI_PRODUIT": SUIVI_PRODUIT_CONFIG,
}

# Fichier CRM (lancé par update_crm.py)
CRM_CONFIG = {
    "SUIVI_CRM": SUIVI_CRM_CONFIG,
}
