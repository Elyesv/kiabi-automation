"""
Configuration pour l'automatisation des fichiers Excel OneDrive.
"""
import os
from pathlib import Path
try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

# Répertoire racine du projet
PROJECT_ROOT = Path(__file__).parent

# Répertoires
LOGS_DIR = PROJECT_ROOT / "logs"
LOGS_DIR.mkdir(exist_ok=True)

# Chemin racine OneDrive (dossier contenant les sous-dossiers de fichiers)
def get_onedrive_path():
    """Retourne le chemin OneDrive configuré."""
    return Path(os.getenv("ONEDRIVE_BASE_PATH", ""))

# Pour compatibilité avec le code existant
ONEDRIVE_BASE_PATH = get_onedrive_path()

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
    "date_sheet": "REPORT",
    "date_cell": "A1",
    "timeout_refresh": 300,
    # Requêtes Power Query à mettre à jour
    "queries": {
        # Requêtes selligent : mettre à jour le numéro de semaine dans le chemin
        "selligent_all": {"type": "selligent"},
        "selligent_all_histo": {"type": "selligent"},
        # Requêtes push : mettre à jour le numéro de semaine dans le chemin
        "push_all": {"type": "selligent"},
        "push_all_histo": {"type": "selligent"},
        # Requêtes piano : mettre à jour les dates start/end (+7 jours)
        "piano_all": {"type": "piano"},
        "piano_all_histo": {"type": "piano"},
    },
}

# Fichier KPIS seul (lancé par update_kpis.py)
KPIS_CONFIG = {
    "SUIVI_KPIS": SUIVI_KPIS_CONFIG,
}

# Fichiers MDR, PMA, PRODUIT (lancés par update_autres.py)
AUTRES_CONFIGS = {
    "SUIVI_MDR": SUIVI_MDR_CONFIG,
    "SUIVI_PMA": SUIVI_PMA_CONFIG,
    "SUIVI_PRODUIT": SUIVI_PRODUIT_CONFIG,
}

# Tous les fichiers (pour clean.py)
FILE_CONFIGS = {
    "SUIVI_KPIS": SUIVI_KPIS_CONFIG,
    "SUIVI_MDR": SUIVI_MDR_CONFIG,
    "SUIVI_PMA": SUIVI_PMA_CONFIG,
    "SUIVI_PRODUIT": SUIVI_PRODUIT_CONFIG,
}

SUIVI_TRAFIC_CONFIG = {
    "folder": "SUIVI_TRAFIC",
    "file_prefix": "SUIVI_TRAFIC",
    "timeout_refresh": 300,
    # Liaisons externes à mettre à jour (vers CRM et KPIS)
    "linked_files": ["SUIVI_CRM", "SUIVI_KPIS"],
    # Requêtes Power Query piano
    "queries": {
        "piano_all": {"type": "piano"},
        "piano_all_histo": {"type": "piano"},
    },
}

# Fichier CRM (lancé par update_crm.py)
CRM_CONFIG = {
    "SUIVI_CRM": SUIVI_CRM_CONFIG,
}

# Fichier TRAFIC (lancé par update_trafic.py)
TRAFIC_CONFIG = {
    "SUIVI_TRAFIC": SUIVI_TRAFIC_CONFIG,
}
