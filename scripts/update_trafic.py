"""
Script d'automatisation de la mise à jour du fichier SUIVI_TRAFIC.

Ce script :
1. Duplique le fichier SUIVI_TRAFIC de la semaine la plus récente
2. Met à jour les liaisons externes vers les nouveaux fichiers CRM et KPIS
3. Met à jour les requêtes Power Query piano (dates +7 jours)
4. Actualise toutes les connexions de données
5. Sauvegarde et ferme
"""
import sys
import re
import shutil
from pathlib import Path
from datetime import datetime, timedelta

sys.path.insert(0, str(Path(__file__).parent.parent))

from config import ONEDRIVE_BASE_PATH, SUIVI_TRAFIC_CONFIG, SUIVI_KPIS_CONFIG, SUIVI_CRM_CONFIG
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


def update_piano_query(excel, query_name: str) -> bool:
    """Met à jour les dates start/end dans une requête piano (+7 jours)."""
    formula = excel.get_query_formula(query_name)
    if formula is None:
        return False

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


def update_external_links(excel, old_week: int, new_week: int, linked_prefixes: list) -> bool:
    """
    Met à jour les liaisons externes en changeant le numéro de semaine.
    """
    links = excel.get_external_links()
    if not links:
        print("  Aucune liaison externe trouvée")
        return True

    print(f"  {len(links)} liaison(s) externe(s) trouvée(s)")
    success = True

    for old_link in links:
        # Vérifier si le lien concerne un des fichiers liés (CRM, KPIS)
        for prefix in linked_prefixes:
            if prefix in old_link:
                # Remplacer _SXX par _S(XX+1)
                pattern = rf'(_S){old_week:02d}'
                replacement = rf'\g<1>{new_week:02d}'
                new_link = re.sub(pattern, replacement, old_link)

                if new_link != old_link:
                    if excel.change_link_path(old_link, new_link):
                        print(f"    {Path(old_link).name} -> {Path(new_link).name}")
                    else:
                        success = False
                break

    return success


def main():
    print("=" * 60)
    print("   MISE A JOUR AUTOMATIQUE - SUIVI_TRAFIC")
    print(f"   {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 60)

    if not ONEDRIVE_BASE_PATH or not ONEDRIVE_BASE_PATH.exists():
        print(f"ERREUR: Chemin OneDrive invalide: {ONEDRIVE_BASE_PATH}")
        return False

    config = SUIVI_TRAFIC_CONFIG
    folder = ONEDRIVE_BASE_PATH / config["folder"]
    prefix = config["file_prefix"]
    ext = config.get("file_ext", ".xlsx")

    if not folder.exists():
        print(f"ERREUR: Dossier introuvable: {folder}")
        return False

    # 1. Trouver le dernier fichier
    print(f"\n[1/6] Recherche du dernier fichier {prefix}_SXX{ext}...")
    source_file, source_week = find_latest_file(folder, prefix, ext)
    if not source_file:
        return False
    next_week = source_week + 1
    print(f"  Trouvé: {source_file.name} (S{source_week:02d} -> S{next_week:02d})")

    # 2. Dupliquer et renommer
    new_name = f"{prefix}_S{next_week:02d}{ext}"
    new_file = folder / new_name
    print(f"\n[2/6] Duplication: {source_file.name} -> {new_name}")

    if new_file.exists():
        print(f"  ATTENTION: {new_name} existe déjà, il sera écrasé")

    shutil.copy2(source_file, new_file)

    # 3-6. Ouvrir et mettre à jour
    excel = None
    success = False

    try:
        excel = ExcelAutomation(visible=True)

        if not excel.open_workbook(new_file):
            print("ERREUR: Impossible d'ouvrir le fichier")
            return False

        # 3. Mettre à jour les liaisons externes
        print(f"\n[3/6] Mise à jour des liaisons externes (CRM, KPIS)...")
        linked_prefixes = config.get("linked_files", [])
        update_external_links(excel, source_week, next_week, linked_prefixes)

        # 4. Mettre à jour les requêtes piano
        print(f"\n[4/6] Mise à jour des requêtes Power Query (piano)...")
        queries = config.get("queries", {})
        for query_name, query_config in queries.items():
            print(f"\n  --- {query_name} ---")
            update_piano_query(excel, query_name)

        # 5. Actualiser les données
        print(f"\n[5/6] Actualisation des données...")
        if not excel.refresh_all_queries(timeout=config["timeout_refresh"]):
            print("ATTENTION: L'actualisation peut ne pas être complète")

        print("  Vérification des connexions...")
        excel.check_connections_status()

        # 6. Sauvegarder et fermer
        print(f"\n[6/6] Sauvegarde et fermeture...")
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
