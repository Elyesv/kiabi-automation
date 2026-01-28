"""
Supprime les fichiers générés (le SXX le plus élevé dans chaque dossier)
pour pouvoir relancer le script sur la semaine précédente.
"""
import sys
import re
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent))

from config import ONEDRIVE_BASE_PATH, FILE_CONFIGS


def main():
    if not ONEDRIVE_BASE_PATH or not ONEDRIVE_BASE_PATH.exists():
        print(f"ERREUR: Chemin OneDrive invalide: {ONEDRIVE_BASE_PATH}")
        return

    for name, config in FILE_CONFIGS.items():
        folder = ONEDRIVE_BASE_PATH / config["folder"]
        prefix = config["file_prefix"]
        ext = config.get("file_ext", ".xlsx")

        if not folder.exists():
            print(f"{name}: dossier introuvable")
            continue

        # Trouver le fichier avec le SXX le plus élevé
        best_file = None
        best_week = -1
        for f in folder.glob(f"{prefix}_S*{ext}"):
            match = re.search(rf'{re.escape(prefix)}_S(\d+)', f.stem)
            if match:
                week = int(match.group(1))
                if week > best_week:
                    best_week = week
                    best_file = f

        if best_file:
            best_file.unlink()
            print(f"{name}: supprimé {best_file.name}")
        else:
            print(f"{name}: aucun fichier trouvé")


if __name__ == "__main__":
    main()
