"""
Configuration pour l'automatisation des fichiers Excel SharePoint.
"""
import os
from pathlib import Path
from dotenv import load_dotenv

# Charger les variables d'environnement
load_dotenv()

# Répertoire racine du projet
PROJECT_ROOT = Path(__file__).parent

# Répertoires
TEMP_DIR = PROJECT_ROOT / "temp"
LOGS_DIR = PROJECT_ROOT / "logs"

# Créer les répertoires s'ils n'existent pas
TEMP_DIR.mkdir(exist_ok=True)
LOGS_DIR.mkdir(exist_ok=True)

# Configuration SharePoint
SHAREPOINT_SITE_URL = os.getenv(
    "SHAREPOINT_SITE_URL",
    "https://kiabi.sharepoint.com/sites/KiabiDigital"
)
SHAREPOINT_EMAIL = os.getenv("SHAREPOINT_EMAIL", "")
SHAREPOINT_PASSWORD = os.getenv("SHAREPOINT_PASSWORD", "")

# Chemin relatif du dossier Documents partagés sur SharePoint
SHAREPOINT_DOC_LIBRARY = "Documents partages"

# Configuration SUIVI_KPIS
SUIVI_KPIS_CONFIG = {
    "sharepoint_folder": "/DIGITAL ANALYTICS/1. Organisation/1. EQUIPE/Back-up/Analyse mensuelle - Process",
    "file_pattern": "SUIVI_KPIS*.xlsx",
    "verification_sheet": "SUIVI_JOUR",
    "timeout_refresh": 300,  # Timeout en secondes pour l'actualisation des requêtes
}

# Mapping des configurations par type de fichier (pour extensions futures)
FILE_CONFIGS = {
    "SUIVI_KPIS": SUIVI_KPIS_CONFIG,
}
