import os
import subprocess
import pythoncom
import win32com.client
from typing import Optional
from SE_to_PLM.core.services.logging.logger_service import logger

class SolidEdgeConnectionManager:
    """
    Gère la connexion COM à Solid Edge.
    Permet de s'attacher à une instance active ou d'en lancer une nouvelle.
    """
    def __init__(self):
        self._app: Optional[win32com.client.CDispatch] = None

    def connect(self, visible: bool = False) -> Optional[win32com.client.CDispatch]:
        """
        Se connecte à Solid Edge. Tente d'abord de récupérer une instance active.
        """
        pythoncom.CoInitialize()
        try:
            # Tente de récupérer l'instance active
            self._app = win32com.client.GetActiveObject("SolidEdge.Application")
            logger.info("Connexion à l'instance active de Solid Edge réussie.")
        except Exception:
            try:
                # Lance une nouvelle instance
                self._app = win32com.client.Dispatch("SolidEdge.Application")
                self._app.Visible = visible
                self._app.DisplayAlerts = False
                logger.info("Démarrage d'une nouvelle session Solid Edge.")
            except Exception as e:
                logger.error(f"Erreur lors du démarrage de Solid Edge : {e}")
                self._app = None
        
        return self._app

    def get_application(self) -> Optional[win32com.client.CDispatch]:
        """Retourne l'objet Application. Se connecte si nécessaire et vérifie la validité."""
        if self._app:
            try:
                # Vérification de base pour s'assurer que le proxy est toujours valide
                # (notamment en cas de changement de thread ou fermeture de SE)
                _ = self._app.Visible
                return self._app
            except Exception:
                logger.warning("Connexion Solid Edge perdue ou invalide. Reconnexion...")
                self._app = None
        
        return self.connect()

    def close_all_documents(self):
        """Ferme tous les documents ouverts dans Solid Edge sans sauvegarder."""
        if not self._app:
            return
        try:
            while self._app.Documents.Count > 0:
                self._app.Documents.Item(1).Close(False)
        except Exception as e:
            logger.warning(f"Erreur lors de la fermeture des documents : {e}")

    def quit(self):
        """Ferme Solid Edge proprement, avec recours au forçage si nécessaire."""
        pythoncom.CoInitialize()
        if self._app:
            try:
                logger.info("Tentative de fermeture propre de Solid Edge...")
                self.close_all_documents()
                self._app.Quit()
                self._app = None
                logger.success("Solid Edge fermé proprement via COM.")
            except Exception as e:
                logger.warning(f"Erreur lors de la fermeture via COM : {e}. Tentative de forçage...")
                self._app = None
        
        # Recours à Taskkill pour garantir la fermeture du processus
        try:
            logger.info("Vérification des processus Edge.exe restants...")
            subprocess.run(["taskkill", "/F", "/IM", "Edge.exe", "/T"], 
                           stdout=subprocess.DEVNULL, 
                           stderr=subprocess.DEVNULL, 
                           check=False)
            logger.info("Processus Solid Edge (Edge.exe) terminés.")
        except Exception as e:
            logger.error(f"Impossible de terminer le processus Solid Edge : {e}")

    def reset_connection(self):
        """Réinitialise la connexion interne sans fermer l'application."""
        self._app = None
        pythoncom.CoUninitialize()
        pythoncom.CoInitialize()

    def __del__(self):
        try:
            pythoncom.CoUninitialize()
        except:
            pass

# Instance globale
connection_manager = SolidEdgeConnectionManager()
