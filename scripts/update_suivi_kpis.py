"""
Script d'automatisation de la mise à jour du fichier SUIVI_KPIS.

Ce script :
1. Télécharge le fichier SUIVI_KPIS depuis SharePoint
2. Ouvre le fichier dans Excel
3. Actualise toutes les requêtes Power Query
4. Vérifie la présence des données de la veille
5. Met à jour les liaisons externes si nécessaire
6. Sauvegarde et uploade le fichier vers SharePoint
"""
import sys
from pathlib import Path
from datetime import datetime

# Ajouter le répertoire parent au path pour les imports
sys.path.insert(0, str(Path(__file__).parent.parent))

from config import (
    SHAREPOINT_SITE_URL,
    SHAREPOINT_EMAIL,
    SHAREPOINT_PASSWORD,
    SHAREPOINT_DOC_LIBRARY,
    SUIVI_KPIS_CONFIG,
    TEMP_DIR,
    LOGS_DIR,
)
from src.sharepoint_client import SharePointClient
from src.excel_automation import ExcelAutomation


def setup_logging():
    """Configure le logging dans un fichier."""
    log_file = LOGS_DIR / f"suivi_kpis_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
    return log_file


def print_header():
    """Affiche l'en-tête du script."""
    print("=" * 60)
    print("   MISE A JOUR AUTOMATIQUE - SUIVI_KPIS")
    print(f"   {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 60)
    print()


def validate_config():
    """Valide la configuration."""
    errors = []

    if not SHAREPOINT_EMAIL:
        errors.append("SHAREPOINT_EMAIL non configuré dans .env")
    if not SHAREPOINT_PASSWORD:
        errors.append("SHAREPOINT_PASSWORD non configuré dans .env")
    if not SHAREPOINT_SITE_URL:
        errors.append("SHAREPOINT_SITE_URL non configuré dans .env")

    if errors:
        print("ERREURS DE CONFIGURATION:")
        for error in errors:
            print(f"  - {error}")
        print("\nVeuillez configurer le fichier .env (voir .env.example)")
        return False

    return True


def main():
    """Fonction principale."""
    print_header()

    # Valider la configuration
    if not validate_config():
        return False

    excel = None
    success = False

    try:
        # Étape 1: Connexion à SharePoint
        print("[1/7] Connexion à SharePoint...")
        sp_client = SharePointClient(
            SHAREPOINT_SITE_URL,
            SHAREPOINT_EMAIL,
            SHAREPOINT_PASSWORD
        )

        if not sp_client.test_connection():
            print("ERREUR: Impossible de se connecter à SharePoint")
            return False

        # Étape 2: Recherche du fichier SUIVI_KPIS
        print(f"\n[2/7] Recherche du fichier {SUIVI_KPIS_CONFIG['file_pattern']}...")
        file_info = sp_client.find_file(
            SUIVI_KPIS_CONFIG["sharepoint_folder"],
            SUIVI_KPIS_CONFIG["file_pattern"],
            SHAREPOINT_DOC_LIBRARY
        )

        if not file_info:
            print(f"ERREUR: Fichier non trouvé dans {SUIVI_KPIS_CONFIG['sharepoint_folder']}")
            print("\nFichiers disponibles dans ce dossier:")
            files = sp_client.list_files(
                SUIVI_KPIS_CONFIG["sharepoint_folder"],
                SHAREPOINT_DOC_LIBRARY
            )
            for f in files:
                print(f"  - {f['name']}")
            return False

        print(f"  Fichier trouvé: {file_info['name']}")

        # Étape 3: Téléchargement du fichier
        print(f"\n[3/7] Téléchargement du fichier...")
        local_file = TEMP_DIR / file_info["name"]
        remote_path = f"{SUIVI_KPIS_CONFIG['sharepoint_folder']}/{file_info['name']}"

        if not sp_client.download_file(remote_path, local_file, SHAREPOINT_DOC_LIBRARY):
            print("ERREUR: Impossible de télécharger le fichier")
            return False

        # Étape 4: Ouverture dans Excel
        print(f"\n[4/7] Ouverture du fichier dans Excel...")
        excel = ExcelAutomation(visible=True)

        if not excel.open_workbook(local_file):
            print("ERREUR: Impossible d'ouvrir le fichier Excel")
            return False

        print(f"  Feuilles disponibles: {', '.join(excel.get_sheet_names())}")

        # Étape 5: Actualisation des requêtes Power Query
        print(f"\n[5/7] Actualisation des requêtes Power Query...")
        if not excel.refresh_all_queries(timeout=SUIVI_KPIS_CONFIG["timeout_refresh"]):
            print("ATTENTION: L'actualisation peut ne pas être complète")

        # Étape 6: Vérification des données
        print(f"\n[6/7] Vérification des données...")

        # Mise à jour des liaisons externes
        excel.update_external_links()

        # Vérification de l'onglet SUIVI_JOUR
        verification_sheet = SUIVI_KPIS_CONFIG.get("verification_sheet")
        if verification_sheet and verification_sheet in excel.get_sheet_names():
            check_result = excel.check_sheet_data(verification_sheet, check_yesterday=True)
            if not check_result.get("yesterday_data", True):
                print(f"\n  ATTENTION: Les données de la veille ne semblent pas présentes")
                print("  Veuillez vérifier manuellement l'onglet SUIVI_JOUR")

        # Sauvegarde
        print("\n  Sauvegarde du fichier...")
        if not excel.save():
            print("ERREUR: Impossible de sauvegarder le fichier")
            return False

        # Fermeture d'Excel
        excel.close(save=False)  # Déjà sauvegardé

        # Étape 7: Upload vers SharePoint
        print(f"\n[7/7] Upload vers SharePoint...")
        if not sp_client.upload_file(
            local_file,
            SUIVI_KPIS_CONFIG["sharepoint_folder"],
            file_info["name"],
            SHAREPOINT_DOC_LIBRARY
        ):
            print("ERREUR: Impossible d'uploader le fichier")
            return False

        success = True
        print("\n" + "=" * 60)
        print("   MISE A JOUR TERMINEE AVEC SUCCES")
        print("=" * 60)

    except Exception as e:
        print(f"\nERREUR INATTENDUE: {e}")
        import traceback
        traceback.print_exc()

    finally:
        # Nettoyage
        if excel:
            try:
                excel.quit()
            except:
                pass

    return success


if __name__ == "__main__":
    success = main()
    print("\nAppuyez sur Entrée pour fermer...")
    input()
    sys.exit(0 if success else 1)
