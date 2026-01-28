"""
Script d'automatisation de la mise à jour du fichier SUIVI_KPIS uniquement.
"""
import sys
import re
import shutil
from pathlib import Path
from datetime import datetime, timedelta

sys.path.insert(0, str(Path(__file__).parent.parent))

from config import ONEDRIVE_BASE_PATH, KPIS_CONFIG
from src.excel_automation import ExcelAutomation


def find_latest_file(folder: Path, prefix: str, ext: str = ".xlsx"):
    """Trouve le fichier avec le numéro de semaine le plus élevé."""
    pattern = f"{prefix}_S*{ext}"
    matches = list(folder.glob(pattern))

    if not matches:
        print(f"  ERREUR: Aucun fichier trouvé pour '{pattern}' dans {folder}")
        return None, None

    best_file = None
    best_week = -1
    for f in matches:
        match = re.search(rf'{re.escape(prefix)}_S(\d+)', f.stem)
        if match:
            week = int(match.group(1))
            if week > best_week:
                best_week = week
                best_file = f

    return best_file, best_week


def process_file(name: str, config: dict, excel: ExcelAutomation) -> bool:
    """Traite un fichier."""
    print(f"\n{'=' * 60}")
    print(f"   {name}")
    print(f"{'=' * 60}")

    folder = ONEDRIVE_BASE_PATH / config["folder"]
    prefix = config["file_prefix"]
    ext = config.get("file_ext", ".xlsx")

    if not folder.exists():
        print(f"  ERREUR: Dossier introuvable: {folder}")
        return False

    print(f"\n  [1/5] Recherche du dernier fichier...")
    source_file, source_week = find_latest_file(folder, prefix, ext)
    if not source_file:
        return False
    next_week = source_week + 1
    print(f"  Trouvé: {source_file.name} (S{source_week:02d} -> S{next_week:02d})")

    new_name = f"{prefix}_S{next_week:02d}{ext}"
    new_file = folder / new_name
    print(f"\n  [2/5] Duplication: {source_file.name} -> {new_name}")

    if new_file.exists():
        print(f"  ATTENTION: {new_name} existe déjà, il sera écrasé")

    shutil.copy2(source_file, new_file)

    try:
        print(f"\n  [3/5] Mise à jour de la date...")

        if not excel.open_workbook(new_file):
            return False

        sheet = config["date_sheet"]
        cell = config["date_cell"]
        current_date = excel.read_cell(sheet, cell)

        if current_date is None:
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

        print(f"\n  [4/5] Actualisation des données...")
        if not excel.refresh_all_queries(timeout=config["timeout_refresh"]):
            print("  ATTENTION: L'actualisation peut ne pas être complète")

        excel.check_connections_status()

        print(f"\n  [5/5] Sauvegarde...")
        if not excel.save():
            return False

        excel.close(save=False)
        print(f"  OK - {new_name} traité avec succès")
        return True

    except Exception as e:
        print(f"  ERREUR: {e}")
        import traceback
        traceback.print_exc()
        try:
            excel.close(save=False)
        except:
            pass
        return False


def main():
    print("=" * 60)
    print("   MISE A JOUR - SUIVI_KPIS")
    print(f"   {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 60)

    if not ONEDRIVE_BASE_PATH or not ONEDRIVE_BASE_PATH.exists():
        print(f"ERREUR: Chemin OneDrive invalide: {ONEDRIVE_BASE_PATH}")
        return False

    excel = None
    results = {}
    try:
        excel = ExcelAutomation(visible=True)
        for name, config in KPIS_CONFIG.items():
            results[name] = process_file(name, config, excel)
    except Exception as e:
        print(f"\nERREUR: {e}")
        return False
    finally:
        if excel:
            try:
                excel.quit()
            except:
                pass

    all_ok = all(results.values())
    print("\n" + "=" * 60)
    if all_ok:
        print("   SUIVI_KPIS: OK")
    else:
        print("   SUIVI_KPIS: ERREUR")
    print("=" * 60)
    return all_ok


if __name__ == "__main__":
    success = main()
    print("\nAppuyez sur Entrée pour fermer...")
    input()
    sys.exit(0 if success else 1)
