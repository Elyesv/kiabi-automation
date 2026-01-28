"""
Lanceur principal pour l'automatisation des fichiers SUIVI.
Demande le chemin OneDrive au premier lancement, puis lance toutes les mises à jour.
"""
import os
import sys
from pathlib import Path

# Pour PyInstaller : obtenir le bon chemin
if getattr(sys, 'frozen', False):
    BASE_DIR = Path(sys.executable).parent
else:
    BASE_DIR = Path(__file__).parent

CONFIG_FILE = BASE_DIR / "config_path.txt"


def get_onedrive_path(force_ask=False):
    """Récupère ou demande le chemin OneDrive."""
    if not force_ask and CONFIG_FILE.exists():
        path = CONFIG_FILE.read_text().strip()
        if path and Path(path).exists():
            return path

    print()
    print("=" * 60)
    print("   CONFIGURATION DU CHEMIN ONEDRIVE")
    print("=" * 60)
    print()
    print("Entrez le chemin du dossier OneDrive contenant les fichiers SUIVI.")
    print("Exemple: C:\\Users\\NomUtilisateur\\OneDrive - Kiabi\\MonDossier")
    print()

    while True:
        path = input("Chemin: ").strip().strip('"')
        if Path(path).exists():
            CONFIG_FILE.write_text(path)
            print(f"\nChemin enregistré: {path}")
            return path
        else:
            print(f"ERREUR: Le dossier '{path}' n'existe pas. Réessayez.")


def main():
    print()
    print("=" * 60)
    print("   AUTOMATISATION FICHIERS SUIVI KIABI")
    print("=" * 60)

    force_config = "--config" in sys.argv

    onedrive_path = get_onedrive_path(force_ask=force_config)
    os.environ["ONEDRIVE_BASE_PATH"] = onedrive_path

    print(f"\nDossier OneDrive: {onedrive_path}")
    print("\n(Pour changer le chemin, relancez avec: Automatisation_SUIVI.exe --config)")

    sys.path.insert(0, str(BASE_DIR))

    import importlib
    import config
    importlib.reload(config)

    print("\n" + "=" * 60)
    print("   LANCEMENT DES MISES A JOUR")
    print("=" * 60)

    results = {}

    # 1. KPIS seul
    print("\n>>> [1/4] Mise à jour SUIVI_KPIS...")
    try:
        from scripts.update_kpis import main as run_kpis
        results["KPIS"] = run_kpis()
    except Exception as e:
        print(f"ERREUR: {e}")
        results["KPIS"] = False

    # 2. MDR, PMA, PRODUIT
    print("\n>>> [2/4] Mise à jour MDR, PMA, PRODUIT...")
    try:
        from scripts.update_autres import main as run_autres
        results["MDR/PMA/PRODUIT"] = run_autres()
    except Exception as e:
        print(f"ERREUR: {e}")
        results["MDR/PMA/PRODUIT"] = False

    # 3. CRM
    print("\n>>> [3/4] Mise à jour CRM...")
    try:
        from scripts.update_crm import main as run_crm
        results["CRM"] = run_crm()
    except Exception as e:
        print(f"ERREUR: {e}")
        results["CRM"] = False

    # 4. TRAFIC
    print("\n>>> [4/4] Mise à jour TRAFIC...")
    try:
        from scripts.update_trafic import main as run_trafic
        results["TRAFIC"] = run_trafic()
    except Exception as e:
        print(f"ERREUR: {e}")
        results["TRAFIC"] = False

    # Résumé final
    print("\n" + "=" * 60)
    print("   RÉSUMÉ FINAL")
    print("=" * 60)
    for name, success in results.items():
        status = "OK" if success else "ERREUR"
        print(f"  {name}: {status}")

    all_ok = all(results.values())
    if all_ok:
        print("\n  Toutes les mises à jour ont été effectuées avec succès!")
    else:
        print("\n  Certaines mises à jour ont échoué. Vérifiez les messages ci-dessus.")

    print("\n" + "=" * 60)
    input("\nAppuyez sur Entrée pour fermer...")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"\nERREUR: {e}")
        import traceback
        traceback.print_exc()
        input("\nAppuyez sur Entrée pour fermer...")
