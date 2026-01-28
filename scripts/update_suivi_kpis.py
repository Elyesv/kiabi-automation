"""
Script d'automatisation de la mise à jour du fichier SUIVI_KPIS.

Ce script :
1. Calcule la semaine courante et la semaine précédente
2. Duplique le fichier SUIVI_KPIS de la semaine précédente
3. Met à jour la date dans REPORT_HEBDO (A1 + 7 jours)
4. Actualise toutes les connexions de données
5. Vérifie que les connexions sont OK
6. Sauvegarde et ferme le fichier
"""
import sys
import shutil
import glob
from pathlib import Path
from datetime import datetime, timedelta

# Ajouter le répertoire parent au path pour les imports
sys.path.insert(0, str(Path(__file__).parent.parent))

from config import (
    ONEDRIVE_BASE_PATH,
    SUIVI_KPIS_CONFIG,
    LOGS_DIR,
)
from src.excel_automation import ExcelAutomation


def print_header():
    """Affiche l'en-tête du script."""
    print("=" * 60)
    print("   MISE A JOUR AUTOMATIQUE - SUIVI_KPIS")
    print(f"   {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 60)
    print()


def get_week_numbers():
    """
    Calcule le numéro de semaine courante (ISO) et le précédent.

    Returns:
        Tuple (semaine_precedente, semaine_courante)
    """
    today = datetime.now()
    current_week = today.isocalendar()[1]
    previous_week = (today - timedelta(weeks=1)).isocalendar()[1]
    return previous_week, current_week


def find_source_file(folder: Path, prefix: str, week_num: int) -> Path:
    """
    Trouve le fichier source de la semaine précédente.

    Args:
        folder: Dossier contenant les fichiers
        prefix: Préfixe du fichier (ex: "SUIVI_KPIS")
        week_num: Numéro de semaine à chercher

    Returns:
        Path du fichier trouvé ou None
    """
    pattern = f"{prefix}_S{week_num:02d}*.xlsx"
    matches = list(folder.glob(pattern))

    if not matches:
        print(f"ERREUR: Aucun fichier trouvé pour le pattern '{pattern}' dans {folder}")
        print("\nFichiers disponibles:")
        for f in sorted(folder.glob(f"{prefix}*.xlsx")):
            print(f"  - {f.name}")
        return None

    # Prendre le plus récent si plusieurs correspondances
    source = max(matches, key=lambda f: f.stat().st_mtime)
    return source


def validate_config():
    """Valide la configuration."""
    if not ONEDRIVE_BASE_PATH or not ONEDRIVE_BASE_PATH.exists():
        print(f"ERREUR: Le chemin OneDrive n'existe pas: {ONEDRIVE_BASE_PATH}")
        print("Veuillez configurer ONEDRIVE_BASE_PATH dans le fichier .env")
        return False

    folder = ONEDRIVE_BASE_PATH / SUIVI_KPIS_CONFIG["folder"]
    if not folder.exists():
        print(f"ERREUR: Le dossier SUIVI_KPI n'existe pas: {folder}")
        return False

    return True


def main():
    """Fonction principale."""
    print_header()

    # Valider la configuration
    if not validate_config():
        return False

    config = SUIVI_KPIS_CONFIG
    folder = ONEDRIVE_BASE_PATH / config["folder"]
    prefix = config["file_prefix"]

    # Étape 1: Calculer les semaines
    prev_week, curr_week = get_week_numbers()
    print(f"[1/6] Calcul des semaines...")
    print(f"  Semaine précédente: S{prev_week:02d}")
    print(f"  Semaine courante:   S{curr_week:02d}")

    # Étape 2: Trouver le fichier source
    print(f"\n[2/6] Recherche du fichier source...")
    source_file = find_source_file(folder, prefix, prev_week)
    if not source_file:
        return False
    print(f"  Fichier trouvé: {source_file.name}")

    # Étape 3: Dupliquer et renommer
    new_name = f"{prefix}_S{curr_week:02d}.xlsx"
    new_file = folder / new_name
    print(f"\n[3/6] Duplication du fichier...")
    print(f"  {source_file.name} -> {new_name}")

    if new_file.exists():
        print(f"  ATTENTION: {new_name} existe déjà, il sera écrasé")

    shutil.copy2(source_file, new_file)
    print(f"  Fichier créé: {new_file}")

    # Étape 4: Ouvrir dans Excel et mettre à jour la date
    excel = None
    success = False

    try:
        print(f"\n[4/6] Ouverture dans Excel et mise à jour de la date...")
        excel = ExcelAutomation(visible=True)

        if not excel.open_workbook(new_file):
            print("ERREUR: Impossible d'ouvrir le fichier Excel")
            return False

        print(f"  Feuilles: {', '.join(excel.get_sheet_names())}")

        # Lire la date actuelle en A1 de REPORT_HEBDO
        sheet = config["date_sheet"]
        cell = config["date_cell"]
        current_date = excel.read_cell(sheet, cell)

        if current_date is None:
            print(f"ERREUR: Impossible de lire la date dans {sheet}!{cell}")
            return False

        # Calculer la nouvelle date (+7 jours)
        if isinstance(current_date, datetime):
            new_date = current_date + timedelta(days=7)
        else:
            print(f"ATTENTION: La valeur en {cell} n'est pas une date: {current_date}")
            print("  Tentative de conversion...")
            try:
                current_date = datetime.strptime(str(current_date), "%Y-%m-%d %H:%M:%S")
                new_date = current_date + timedelta(days=7)
            except ValueError:
                print(f"ERREUR: Impossible de convertir '{current_date}' en date")
                return False

        print(f"  Date actuelle: {current_date.strftime('%d/%m/%Y')}")
        print(f"  Nouvelle date: {new_date.strftime('%d/%m/%Y')}")

        if not excel.write_cell(sheet, cell, new_date):
            print("ERREUR: Impossible d'écrire la nouvelle date")
            return False

        # Étape 5: Actualiser toutes les connexions
        print(f"\n[5/6] Actualisation des données (Données -> Actualiser tout)...")
        if not excel.refresh_all_queries(timeout=config["timeout_refresh"]):
            print("ATTENTION: L'actualisation peut ne pas être complète")

        # Vérifier l'état des connexions
        print("\n  Vérification des connexions (Requêtes et Connexions)...")
        excel.check_connections_status()

        # Étape 6: Sauvegarder et fermer
        print(f"\n[6/6] Sauvegarde et fermeture...")
        if not excel.save():
            print("ERREUR: Impossible de sauvegarder")
            return False

        excel.close(save=False)  # Déjà sauvegardé
        success = True

        print("\n" + "=" * 60)
        print("   MISE A JOUR TERMINEE AVEC SUCCES")
        print(f"   Fichier: {new_name}")
        print(f"   Date mise à jour: {new_date.strftime('%d/%m/%Y')}")
        print("=" * 60)

    except Exception as e:
        print(f"\nERREUR INATTENDUE: {e}")
        import traceback
        traceback.print_exc()

    finally:
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
