import pythoncom
from PyQt5.QtCore import QThread, pyqtSignal
from SE_to_PLM.core.use_cases.generate_plm_export import generate_plm_export_use_case
from SE_to_PLM.core.use_cases.batch_export import batch_export_use_case
from SE_to_PLM.core.use_cases.unified_multi_asm_export import unified_multi_asm_export_use_case
from SE_to_PLM.core.services.logging.logger_service import logger, LogLevel
from SE_to_PLM.core.services.cache.metadata_cache import metadata_cache
from SE_to_PLM.infrastructure.solid_edge.connection_manager import connection_manager

class ExtractionThread(QThread):
    """
    Thread worker pour exécuter l'extraction sans geler l'interface.
    Gère les modes Assemblage, Lot, et Multi-ASM Unifié.
    """
    # Signaux pour communiquer avec la fenêtre principale
    progress_signal = pyqtSignal(int, int, str)  # valeur, max, message
    log_signal = pyqtSignal(str, str)            # message, nom_niveau
    finished_signal = pyqtSignal(bool, str)      # succès, message_final

    def __init__(self, input_path: str, output_dir: str, output_name: str, 
                 dft_folder: str = None, search_mode: str = "les_deux", 
                 mode: str = "assembly", recursive: bool = False,
                 is_folder_source: bool = False, output_mode: str = "single",
                 chunk_index: int = 0, chunk_size: int = -1):
        super().__init__()
        self.input_path = input_path
        self.output_dir = output_dir
        self.output_name = output_name
        self.dft_folder = dft_folder
        self.search_mode = search_mode
        self.mode = mode # "assembly", "batch", ou "multi_asm"
        self.recursive = recursive
        self.is_folder_source = is_folder_source
        self.output_mode = output_mode # "single" ou "multiple"
        self.chunk_index = chunk_index
        self.chunk_size = chunk_size
        self._is_cancelled = False

    def cancel(self):
        self._is_cancelled = True

    def _check_cancelled(self) -> bool:
        return self._is_cancelled

    def _on_log(self, message: str, level: LogLevel):
        self.log_signal.emit(message, level.value)

    def run(self):
        # Initialise COM pour ce thread
        pythoncom.CoInitialize()
        
        # Enregistre un callback local pour rediriger les logs vers l'UI
        logger.register_callback(self._on_log)
        
        # Vider le cache des métadonnées pour repartir sur une base propre
        metadata_cache.clear()
        
        try:
            if self.mode == "assembly":
                generate_plm_export_use_case.execute(
                    input_file=self.input_path,
                    output_dir=self.output_dir,
                    output_name=self.output_name,
                    dft_folder=self.dft_folder,
                    search_mode=self.search_mode,
                    progress_callback=self.progress_signal.emit,
                    is_cancelled=self._check_cancelled
                )
            elif self.mode == "batch":
                batch_export_use_case.execute(
                    input_dir=self.input_path,
                    output_dir=self.output_dir,
                    output_name=self.output_name,
                    recursive=self.recursive,
                    dft_folder=self.dft_folder,
                    search_mode=self.search_mode,
                    progress_callback=self.progress_signal.emit,
                    is_cancelled=self._check_cancelled
                )
            elif self.mode == "multi_asm":
                unified_multi_asm_export_use_case.execute(
                    input_path=self.input_path,
                    output_dir=self.output_dir,
                    output_name=self.output_name,
                    is_folder_source=self.is_folder_source,
                    output_mode=self.output_mode,
                    dft_folder=self.dft_folder,
                    search_mode=self.search_mode,
                    progress_callback=self.progress_signal.emit,
                    is_cancelled=self._check_cancelled,
                    chunk_size=self.chunk_size,
                    chunk_index=self.chunk_index
                )
            
            if self._is_cancelled:
                self.finished_signal.emit(False, "Extraction annulée par l'utilisateur.")
            else:
                self.finished_signal.emit(True, "Extraction terminée avec succès.")
                
        except Exception as e:
            self.finished_signal.emit(False, f"Erreur : {str(e)}")
        finally:
            # Nettoyage systématique
            try:
                connection_manager.close_all_documents()
            except:
                pass
                
            # RÉINITIALISATION CRUCIALE : On force la remise à zéro de la référence globale
            # pour que le prochain thread d'extraction ne récupère pas un proxy COM "mort"
            connection_manager._app = None
            
            logger.unregister_callback(self._on_log)
            pythoncom.CoUninitialize()
