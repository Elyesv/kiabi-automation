"""
Client SharePoint pour télécharger et uploader des fichiers.
"""
import fnmatch
from pathlib import Path
from typing import List, Optional

from office365.runtime.auth.user_credential import UserCredential
from office365.sharepoint.client_context import ClientContext
from office365.sharepoint.files.file import File


class SharePointClient:
    """Client pour interagir avec SharePoint."""

    def __init__(self, site_url: str, email: str, password: str):
        """
        Initialise la connexion à SharePoint.

        Args:
            site_url: URL du site SharePoint (ex: https://kiabi.sharepoint.com/sites/KiabiDigital)
            email: Email de l'utilisateur
            password: Mot de passe de l'utilisateur
        """
        self.site_url = site_url
        self.credentials = UserCredential(email, password)
        self.ctx = ClientContext(site_url).with_credentials(self.credentials)

    def test_connection(self) -> bool:
        """
        Teste la connexion à SharePoint.

        Returns:
            True si la connexion est réussie, False sinon
        """
        try:
            web = self.ctx.web
            self.ctx.load(web)
            self.ctx.execute_query()
            print(f"Connecté à SharePoint: {web.properties['Title']}")
            return True
        except Exception as e:
            print(f"Erreur de connexion SharePoint: {e}")
            return False

    def list_files(self, folder_path: str, doc_library: str = "Documents partages") -> List[dict]:
        """
        Liste les fichiers dans un dossier SharePoint.

        Args:
            folder_path: Chemin relatif du dossier (ex: /DIGITAL ANALYTICS/...)
            doc_library: Nom de la bibliothèque de documents

        Returns:
            Liste des fichiers avec leurs métadonnées
        """
        full_path = f"{doc_library}{folder_path}"
        folder = self.ctx.web.get_folder_by_server_relative_url(full_path)
        files = folder.files
        self.ctx.load(files)
        self.ctx.execute_query()

        return [
            {
                "name": f.properties["Name"],
                "url": f.properties["ServerRelativeUrl"],
                "size": f.properties.get("Length", 0),
                "modified": f.properties.get("TimeLastModified", ""),
            }
            for f in files
        ]

    def find_file(
        self,
        folder_path: str,
        pattern: str,
        doc_library: str = "Documents partages"
    ) -> Optional[dict]:
        """
        Trouve un fichier correspondant à un pattern dans un dossier.

        Args:
            folder_path: Chemin relatif du dossier
            pattern: Pattern de nom de fichier (ex: SUIVI_KPIS*.xlsx)
            doc_library: Nom de la bibliothèque de documents

        Returns:
            Dictionnaire avec les infos du fichier ou None si non trouvé
        """
        files = self.list_files(folder_path, doc_library)
        matching_files = [f for f in files if fnmatch.fnmatch(f["name"], pattern)]

        if not matching_files:
            return None

        # Retourner le fichier le plus récent si plusieurs correspondent
        matching_files.sort(key=lambda x: x.get("modified", ""), reverse=True)
        return matching_files[0]

    def download_file(
        self,
        remote_path: str,
        local_path: Path,
        doc_library: str = "Documents partages"
    ) -> bool:
        """
        Télécharge un fichier depuis SharePoint.

        Args:
            remote_path: Chemin relatif du fichier sur SharePoint
            local_path: Chemin local où sauvegarder le fichier
            doc_library: Nom de la bibliothèque de documents

        Returns:
            True si le téléchargement est réussi, False sinon
        """
        try:
            full_path = f"{doc_library}{remote_path}"
            file = self.ctx.web.get_file_by_server_relative_url(full_path)
            self.ctx.load(file)
            self.ctx.execute_query()

            with open(local_path, "wb") as f:
                file.download(f).execute_query()

            print(f"Fichier téléchargé: {local_path}")
            return True

        except Exception as e:
            print(f"Erreur de téléchargement: {e}")
            return False

    def upload_file(
        self,
        local_path: Path,
        remote_folder: str,
        remote_filename: Optional[str] = None,
        doc_library: str = "Documents partages"
    ) -> bool:
        """
        Upload un fichier vers SharePoint.

        Args:
            local_path: Chemin local du fichier à uploader
            remote_folder: Dossier de destination sur SharePoint
            remote_filename: Nom du fichier sur SharePoint (utilise le nom local par défaut)
            doc_library: Nom de la bibliothèque de documents

        Returns:
            True si l'upload est réussi, False sinon
        """
        try:
            if remote_filename is None:
                remote_filename = local_path.name

            full_folder_path = f"{doc_library}{remote_folder}"
            target_folder = self.ctx.web.get_folder_by_server_relative_url(full_folder_path)

            with open(local_path, "rb") as f:
                content = f.read()

            target_folder.upload_file(remote_filename, content).execute_query()
            print(f"Fichier uploadé: {remote_folder}/{remote_filename}")
            return True

        except Exception as e:
            print(f"Erreur d'upload: {e}")
            return False
