"""
Module d'automatisation Excel via COM (Windows uniquement).
Permet d'actualiser les requêtes Power Query et de manipuler les classeurs.
"""
import time
from pathlib import Path
from typing import Optional, List
from datetime import datetime, timedelta


class ExcelAutomation:
    """
    Classe pour automatiser Excel via COM automation (pywin32).
    Nécessite Windows et Excel installé.
    """

    def __init__(self, visible: bool = True):
        """
        Initialise une instance Excel.

        Args:
            visible: Si True, Excel sera visible pendant l'exécution
        """
        try:
            import win32com.client
            import pythoncom
        except ImportError:
            raise ImportError(
                "pywin32 n'est pas installé. "
                "Exécutez: pip install pywin32"
            )

        pythoncom.CoInitialize()
        self.excel = win32com.client.Dispatch("Excel.Application")
        self.excel.Visible = visible
        self.excel.DisplayAlerts = False
        self.workbook = None
        self._pythoncom = pythoncom

    def open_workbook(self, file_path: Path) -> bool:
        """
        Ouvre un classeur Excel.

        Args:
            file_path: Chemin du fichier Excel

        Returns:
            True si l'ouverture est réussie
        """
        try:
            self.workbook = self.excel.Workbooks.Open(str(file_path.absolute()))
            print(f"Classeur ouvert: {file_path.name}")
            return True
        except Exception as e:
            print(f"Erreur d'ouverture du classeur: {e}")
            return False

    def refresh_all_queries(self, timeout: int = 300) -> bool:
        """
        Actualise toutes les requêtes Power Query du classeur.

        Args:
            timeout: Timeout en secondes pour l'actualisation

        Returns:
            True si l'actualisation est réussie
        """
        if not self.workbook:
            print("Aucun classeur ouvert")
            return False

        try:
            print("Actualisation des requêtes Power Query...")

            # Actualiser toutes les connexions
            self.workbook.RefreshAll()

            # Attendre que toutes les requêtes soient terminées
            start_time = time.time()
            while True:
                # Vérifier si des requêtes sont encore en cours
                refreshing = False
                for connection in self.workbook.Connections:
                    try:
                        if connection.OLEDBConnection:
                            if connection.OLEDBConnection.Refreshing:
                                refreshing = True
                                break
                    except:
                        pass

                if not refreshing:
                    break

                if time.time() - start_time > timeout:
                    print(f"Timeout après {timeout} secondes")
                    return False

                time.sleep(2)

            # Deuxième actualisation (comme mentionné dans le process)
            print("Deuxième actualisation (sécurité)...")
            self.workbook.RefreshAll()
            time.sleep(5)

            print("Actualisation terminée")
            return True

        except Exception as e:
            print(f"Erreur lors de l'actualisation: {e}")
            return False

    def update_external_links(self) -> bool:
        """
        Met à jour toutes les liaisons externes du classeur.

        Returns:
            True si la mise à jour est réussie
        """
        if not self.workbook:
            print("Aucun classeur ouvert")
            return False

        try:
            links = self.workbook.LinkSources(1)  # 1 = xlExcelLinks
            if links:
                print(f"Mise à jour de {len(links)} liaison(s) externe(s)...")
                for link in links:
                    try:
                        self.workbook.UpdateLink(link, 1)
                        print(f"  - Liaison mise à jour: {link}")
                    except Exception as e:
                        print(f"  - Erreur liaison {link}: {e}")
            else:
                print("Aucune liaison externe trouvée")
            return True
        except Exception as e:
            # Pas de liaisons externes n'est pas une erreur
            if "object required" in str(e).lower():
                print("Aucune liaison externe dans ce classeur")
                return True
            print(f"Erreur lors de la mise à jour des liaisons: {e}")
            return False

    def check_sheet_data(
        self,
        sheet_name: str,
        check_yesterday: bool = True
    ) -> dict:
        """
        Vérifie la présence de données dans une feuille.

        Args:
            sheet_name: Nom de la feuille à vérifier
            check_yesterday: Si True, vérifie la présence de données de la veille

        Returns:
            Dictionnaire avec les résultats de la vérification
        """
        if not self.workbook:
            return {"success": False, "error": "Aucun classeur ouvert"}

        try:
            sheet = self.workbook.Sheets(sheet_name)
            used_range = sheet.UsedRange

            result = {
                "success": True,
                "sheet_name": sheet_name,
                "rows": used_range.Rows.Count,
                "columns": used_range.Columns.Count,
                "has_data": used_range.Rows.Count > 1,
            }

            if check_yesterday:
                # Chercher la date de la veille dans la première colonne
                yesterday = (datetime.now() - timedelta(days=1)).strftime("%Y-%m-%d")
                yesterday_found = False

                # Parcourir les premières lignes pour trouver la date
                for row in range(1, min(used_range.Rows.Count + 1, 100)):
                    cell_value = sheet.Cells(row, 1).Value
                    if cell_value:
                        if isinstance(cell_value, datetime):
                            cell_date = cell_value.strftime("%Y-%m-%d")
                            if cell_date == yesterday:
                                yesterday_found = True
                                break

                result["yesterday_data"] = yesterday_found
                if yesterday_found:
                    print(f"Données de la veille ({yesterday}) trouvées dans {sheet_name}")
                else:
                    print(f"ATTENTION: Données de la veille ({yesterday}) non trouvées dans {sheet_name}")

            return result

        except Exception as e:
            return {"success": False, "error": str(e)}

    def get_sheet_names(self) -> List[str]:
        """
        Retourne la liste des noms de feuilles du classeur.

        Returns:
            Liste des noms de feuilles
        """
        if not self.workbook:
            return []

        return [sheet.Name for sheet in self.workbook.Sheets]

    def save(self) -> bool:
        """
        Sauvegarde le classeur.

        Returns:
            True si la sauvegarde est réussie
        """
        if not self.workbook:
            print("Aucun classeur ouvert")
            return False

        try:
            self.workbook.Save()
            print("Classeur sauvegardé")
            return True
        except Exception as e:
            print(f"Erreur de sauvegarde: {e}")
            return False

    def save_as(self, file_path: Path) -> bool:
        """
        Sauvegarde le classeur sous un nouveau nom.

        Args:
            file_path: Chemin du nouveau fichier

        Returns:
            True si la sauvegarde est réussie
        """
        if not self.workbook:
            print("Aucun classeur ouvert")
            return False

        try:
            self.workbook.SaveAs(str(file_path.absolute()))
            print(f"Classeur sauvegardé sous: {file_path}")
            return True
        except Exception as e:
            print(f"Erreur de sauvegarde: {e}")
            return False

    def close(self, save: bool = True) -> bool:
        """
        Ferme le classeur.

        Args:
            save: Si True, sauvegarde avant de fermer

        Returns:
            True si la fermeture est réussie
        """
        if not self.workbook:
            return True

        try:
            self.workbook.Close(SaveChanges=save)
            self.workbook = None
            print("Classeur fermé")
            return True
        except Exception as e:
            print(f"Erreur de fermeture: {e}")
            return False

    def quit(self):
        """Ferme l'application Excel."""
        try:
            if self.workbook:
                self.close(save=False)
            self.excel.Quit()
            self._pythoncom.CoUninitialize()
            print("Excel fermé")
        except Exception as e:
            print(f"Erreur lors de la fermeture d'Excel: {e}")
