"""
Lanceur principal pour l'automatisation des fichiers SUIVI.
Demande le chemin OneDrive au premier lancement, puis propose un menu.
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


def get_onedrive_path():
    """Récupère ou demande le chemin OneDrive."""
    # Vérifier si déjà configuré
    if CONFIG_FILE.exists():
        path = CONFIG_FILE.read_text().strip()
        if path and Path(path).exists():
            return path

    # Demander le chemin
    print("=" * 60)
    print("   CONFIGURATION INITIALE")
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


def show_menu():
    """Affiche le menu principal."""
    print()
    print("=" * 60)
    print("   AUTOMATISATION FICHIERS SUIVI - MENU")
    print("=" * 60)
    print()
    print("  1. Mettre à jour KPIS, MDR, PMA, PRODUIT (les 4)")
    print("  2. Mettre à jour CRM")
    print("  3. Mettre à jour TRAFIC")
    print("  4. Tout mettre à jour (1 + 2 + 3 dans l'ordre)")
    print()
    print("  5. Nettoyer (supprimer derniers fichiers générés)")
    print("  6. Changer le chemin OneDrive")
    print("  0. Quitter")
    print()
    return input("Votre choix: ").strip()


def main():
    print()
    print("=" * 60)
    print("   AUTOMATISATION FICHIERS SUIVI KIABI")
    print("=" * 60)

    # Obtenir le chemin OneDrive AVANT d'importer les modules
    onedrive_path = get_onedrive_path()
    os.environ["ONEDRIVE_BASE_PATH"] = onedrive_path

    print(f"\nDossier OneDrive: {onedrive_path}")

    # Ajouter le répertoire au path pour les imports
    sys.path.insert(0, str(BASE_DIR))

    while True:
        choice = show_menu()

        if choice == "1":
            # Recharger le module config pour prendre en compte le chemin
            import importlib
            import config
            importlib.reload(config)
            from scripts.update_all import main as run_all
            run_all()

        elif choice == "2":
            import importlib
            import config
            importlib.reload(config)
            from scripts.update_crm import main as run_crm
            run_crm()

        elif choice == "3":
            import importlib
            import config
            importlib.reload(config)
            from scripts.update_trafic import main as run_trafic
            run_trafic()

        elif choice == "4":
            import importlib
            import config
            importlib.reload(config)

            print("\n>>> Mise à jour des 4 fichiers principaux...")
            from scripts.update_all import main as run_all
            run_all()

            print("\n>>> Mise à jour CRM...")
            from scripts.update_crm import main as run_crm
            run_crm()

            print("\n>>> Mise à jour TRAFIC...")
            from scripts.update_trafic import main as run_trafic
            run_trafic()

        elif choice == "5":
            import importlib
            import config
            importlib.reload(config)
            from scripts.clean import main as run_clean
            run_clean()

        elif choice == "6":
            CONFIG_FILE.unlink(missing_ok=True)
            onedrive_path = get_onedrive_path()
            os.environ["ONEDRIVE_BASE_PATH"] = onedrive_path
            print(f"\nNouveau chemin: {onedrive_path}")

        elif choice == "0":
            print("\nAu revoir!")
            break

        else:
            print("\nChoix invalide, réessayez.")

        input("\nAppuyez sur Entrée pour continuer...")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"\nERREUR: {e}")
        import traceback
        traceback.print_exc()
        input("\nAppuyez sur Entrée pour fermer...")
