"""
Script d'automatisation de la mise à jour de tous les fichiers SUIVI.

Traite tous les fichiers configurés dans FILE_CONFIGS :
- SUIVI_KPIS, SUIVI_MDR, SUIVI_PMA, SUIVI_PRODUIT

Pour chaque fichier :
1. Duplique le fichier de la semaine précédente
2. Met à jour la date (+7 jours)
3. Actualise toutes les connexions de données
4. Sauvegarde et ferme
"""
import sys
import shutil
from pathlib import Path
from datetime import datetime, timedelta

# Ajouter le répertoire parent au path pour les imports
sys.path.insert(0, str(Path(__file__).parent.parent))

from config import ONEDRIVE_BASE_PATH, FILE_CONFIGS
from src.excel_automation import ExcelAutomation


def get_week_numbers():
    """Calcule le numéro de semaine courante (ISO) et le précédent."""
    today = datetime.now()
    current_week = today.isocalendar()[1]
    previous_week = (today - timedelta(weeks=1)).isocalendar()[1]
    return previous_week, current_week


def find_source_file(folder: Path, prefix: str, week_num: int) -> Path:
    """Trouve le fichier source de la semaine précédente."""
    pattern = f"{prefix}_S{week_num:02d}*.xlsx"
    matches = list(folder.glob(pattern))

    if not matches:
        print(f"  ERREUR: Aucun fichier trouvé pour '{pattern}' dans {folder}")
        print("  Fichiers disponibles:")
        for f in sorted(folder.glob(f"{prefix}*.xlsx")):
            print(f"    - {f.name}")
        return None

    return max(matches, key=lambda f: f.stat().st_mtime)


def process_file(name: str, config: dict, prev_week: int, curr_week: int) -> bool:
    """
    Traite un fichier : duplication, mise à jour date, actualisation données.

    Returns:
        True si le traitement est réussi
    """
    print(f"\n{'=' * 60}")
    print(f"   {name}")
    print(f"{'=' * 60}")

    folder = ONEDRIVE_BASE_PATH / config["folder"]
    prefix = config["file_prefix"]

    # Vérifier que le dossier existe
    if not folder.exists():
        print(f"  ERREUR: Dossier introuvable: {folder}")
        return False

    # 1. Trouver le fichier source
    print(f"\n  [1/5] Recherche de {prefix}_S{prev_week:02d}...")
    source_file = find_source_file(folder, prefix, prev_week)
    if not source_file:
        return False
    print(f"  Trouvé: {source_file.name}")

    # 2. Dupliquer et renommer
    new_name = f"{prefix}_S{curr_week:02d}.xlsx"
    new_file = folder / new_name
    print(f"\n  [2/5] Duplication: {source_file.name} -> {new_name}")

    if new_file.exists():
        print(f"  ATTENTION: {new_name} existe déjà, il sera écrasé")

    shutil.copy2(source_file, new_file)

    # 3. Ouvrir et mettre à jour la date
    excel = None
    try:
        print(f"\n  [3/5] Mise à jour de la date dans {config['date_sheet']}!{config['date_cell']}...")
        excel = ExcelAutomation(visible=True)

        if not excel.open_workbook(new_file):
            print("  ERREUR: Impossible d'ouvrir le fichier")
            return False

        sheet = config["date_sheet"]
        cell = config["date_cell"]
        current_date = excel.read_cell(sheet, cell)

        if current_date is None:
            print(f"  ERREUR: Impossible de lire {sheet}!{cell}")
            return False

        if isinstance(current_date, datetime):
            new_date = current_date + timedelta(days=7)
        else:
            try:
                current_date = datetime.strptime(str(current_date), "%Y-%m-%d %H:%M:%S")
                new_date = current_date + timedelta(days=7)
            except ValueError:
                print(f"  ERREUR: '{current_date}' n'est pas une date valide")
                return False

        print(f"  {current_date.strftime('%d/%m/%Y')} -> {new_date.strftime('%d/%m/%Y')}")

        if not excel.write_cell(sheet, cell, new_date):
            return False

        # 4. Actualiser les données
        print(f"\n  [4/5] Actualisation des données...")
        if not excel.refresh_all_queries(timeout=config["timeout_refresh"]):
            print("  ATTENTION: L'actualisation peut ne pas être complète")

        print("  Vérification des connexions...")
        excel.check_connections_status()

        # 5. Sauvegarder et fermer
        print(f"\n  [5/5] Sauvegarde et fermeture...")
        if not excel.save():
            print("  ERREUR: Impossible de sauvegarder")
            return False

        excel.close(save=False)
        print(f"  OK - {new_name} traité avec succès")
        return True

    except Exception as e:
        print(f"  ERREUR INATTENDUE: {e}")
        import traceback
        traceback.print_exc()
        return False

    finally:
        if excel:
            try:
                excel.quit()
            except:
                pass


def main():
    """Fonction principale : traite tous les fichiers configurés."""
    print("=" * 60)
    print("   MISE A JOUR AUTOMATIQUE - TOUS LES SUIVIS")
    print(f"   {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 60)

    # Vérifier le chemin OneDrive
    if not ONEDRIVE_BASE_PATH or not ONEDRIVE_BASE_PATH.exists():
        print(f"ERREUR: Chemin OneDrive invalide: {ONEDRIVE_BASE_PATH}")
        print("Configurez ONEDRIVE_BASE_PATH dans le fichier .env")
        return False

    prev_week, curr_week = get_week_numbers()
    print(f"\nSemaine précédente: S{prev_week:02d}")
    print(f"Semaine courante:   S{curr_week:02d}")
    print(f"Fichiers à traiter: {len(FILE_CONFIGS)}")

    results = {}
    for name, config in FILE_CONFIGS.items():
        results[name] = process_file(name, config, prev_week, curr_week)

    # Résumé
    print("\n" + "=" * 60)
    print("   RÉSUMÉ")
    print("=" * 60)
    for name, success in results.items():
        status = "OK" if success else "ERREUR"
        print(f"  {name}: {status}")

    all_ok = all(results.values())
    if all_ok:
        print("\n  Tous les fichiers ont été traités avec succès.")
    else:
        failed = [n for n, s in results.items() if not s]
        print(f"\n  {len(failed)} fichier(s) en erreur: {', '.join(failed)}")

    return all_ok


if __name__ == "__main__":
    success = main()
    print("\nAppuyez sur Entrée pour fermer...")
    input()
    sys.exit(0 if success else 1)
