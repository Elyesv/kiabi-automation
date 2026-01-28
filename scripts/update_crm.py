"""
Script d'automatisation de la mise à jour du fichier SUIVI_CRM.

Ce script :
1. Duplique le fichier SUIVI_CRM de la semaine la plus récente
2. Met à jour les requêtes Power Query :
   - selligent_all / selligent_all_histo : chemin fichier avec numéro de semaine
   - piano_all / piano_all_histo : dates start/end (+7 jours)
3. Actualise toutes les connexions de données
4. Sauvegarde et ferme
"""
import sys
import re
import shutil
from pathlib import Path
from datetime import datetime, timedelta

sys.path.insert(0, str(Path(__file__).parent.parent))

from config import ONEDRIVE_BASE_PATH, SUIVI_CRM_CONFIG
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

    if best_file is None:
        print(f"  ERREUR: Aucun fichier avec un numéro de semaine valide")
        return None, None

    return best_file, best_week


def update_selligent_query(excel, query_name: str, old_week: int, new_week: int) -> bool:
    """
    Met à jour le numéro de semaine dans une requête selligent.
    Ex: 2026_S03 -> 2026_S04
    """
    formula = excel.get_query_formula(query_name)
    if formula is None:
        return False

    # Remplacer le pattern YYYY_SXX par YYYY_S(XX+1)
    # Le pattern peut contenir n'importe quelle année
    pattern = rf'(\d{{4}})_S{old_week:02d}'
    replacement = rf'\1_S{new_week:02d}'

    new_formula = re.sub(pattern, replacement, formula)

    if new_formula == formula:
        print(f"  ATTENTION: Aucun pattern S{old_week:02d} trouvé dans '{query_name}'")
        return False

    # Afficher ce qui a changé
    old_matches = re.findall(pattern, formula)
    for year in old_matches:
        print(f"  {year}_S{old_week:02d} -> {year}_S{new_week:02d}")

    return excel.set_query_formula(query_name, new_formula)


def update_piano_query(excel, query_name: str) -> bool:
    """
    Met à jour les dates start/end dans une requête piano (+7 jours).
    Trouve les dates dans le JSON encodé et ajoute 7 jours.
    """
    formula = excel.get_query_formula(query_name)
    if formula is None:
        return False

    # Trouver toutes les dates au format YYYY-MM-DD dans la formule
    date_pattern = r'(\d{4}-\d{2}-\d{2})'
    dates_found = re.findall(date_pattern, formula)

    if not dates_found:
        print(f"  ATTENTION: Aucune date trouvée dans '{query_name}'")
        return False

    new_formula = formula
    for date_str in dates_found:
        old_date = datetime.strptime(date_str, "%Y-%m-%d")
        new_date = old_date + timedelta(days=7)
        new_date_str = new_date.strftime("%Y-%m-%d")
        new_formula = new_formula.replace(date_str, new_date_str)
        print(f"  {date_str} -> {new_date_str}")

    return excel.set_query_formula(query_name, new_formula)


def main():
    print("=" * 60)
    print("   MISE A JOUR AUTOMATIQUE - SUIVI_CRM")
    print(f"   {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 60)

    if not ONEDRIVE_BASE_PATH or not ONEDRIVE_BASE_PATH.exists():
        print(f"ERREUR: Chemin OneDrive invalide: {ONEDRIVE_BASE_PATH}")
        return False

    config = SUIVI_CRM_CONFIG
    folder = ONEDRIVE_BASE_PATH / config["folder"]
    prefix = config["file_prefix"]
    ext = config.get("file_ext", ".xlsx")

    if not folder.exists():
        print(f"ERREUR: Dossier introuvable: {folder}")
        return False

    # 1. Trouver le dernier fichier
    print(f"\n[1/5] Recherche du dernier fichier {prefix}_SXX{ext}...")
    source_file, source_week = find_latest_file(folder, prefix, ext)
    if not source_file:
        return False
    next_week = source_week + 1
    print(f"  Trouvé: {source_file.name} (S{source_week:02d} -> S{next_week:02d})")

    # 2. Dupliquer et renommer
    new_name = f"{prefix}_S{next_week:02d}{ext}"
    new_file = folder / new_name
    print(f"\n[2/5] Duplication: {source_file.name} -> {new_name}")

    if new_file.exists():
        print(f"  ATTENTION: {new_name} existe déjà, il sera écrasé")

    shutil.copy2(source_file, new_file)

    # 3. Ouvrir et mettre à jour les requêtes
    excel = None
    success = False

    try:
        excel = ExcelAutomation(visible=True)

        if not excel.open_workbook(new_file):
            print("ERREUR: Impossible d'ouvrir le fichier")
            return False

        print(f"\n[3/5] Mise à jour des requêtes Power Query...")
        queries = config.get("queries", {})

        for query_name, query_config in queries.items():
            query_type = query_config["type"]
            print(f"\n  --- {query_name} ({query_type}) ---")

            if query_type == "selligent":
                update_selligent_query(excel, query_name, source_week, next_week)
            elif query_type == "piano":
                update_piano_query(excel, query_name)

        # 4. Actualiser les données
        print(f"\n[4/5] Actualisation des données...")
        if not excel.refresh_all_queries(timeout=config["timeout_refresh"]):
            print("ATTENTION: L'actualisation peut ne pas être complète")

        print("  Vérification des connexions...")
        excel.check_connections_status()

        # 5. Sauvegarder et fermer
        print(f"\n[5/5] Sauvegarde et fermeture...")
        if not excel.save():
            print("ERREUR: Impossible de sauvegarder")
            return False

        excel.close(save=False)
        success = True

        print("\n" + "=" * 60)
        print("   MISE A JOUR TERMINEE AVEC SUCCES")
        print(f"   Fichier: {new_name}")
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
