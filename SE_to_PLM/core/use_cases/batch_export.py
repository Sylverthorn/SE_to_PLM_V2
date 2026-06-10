import os
import time
from typing import List, Dict, Optional, Callable
from pathlib import Path

from SE_to_PLM.core.models.enums import PlmClass, RelationshipType
from SE_to_PLM.core.models.cad_file import CadFile
from SE_to_PLM.core.models.export_row import ExportRow
from SE_to_PLM.core.services.solid_edge.assembly_service import assembly_service
from SE_to_PLM.core.services.solid_edge.metadata_service import metadata_service
from SE_to_PLM.core.services.indexing.dft_indexer import dft_indexer
from SE_to_PLM.core.services.export.excel_exporter import excel_exporter
from SE_to_PLM.core.services.plm.revision_service import revision_service
from SE_to_PLM.core.services.logging.logger_service import logger

class BatchExportUseCase:
    """
    Orchestrates the batch export of standalone part/sheetmetal files in a directory.
    """

    def __init__(self):
        self._order_counter = 1

    def _get_attachment_suffix(self, plm_class: PlmClass) -> str:
        if plm_class == PlmClass.SUB_ASSY: return "(ASM)"
        if plm_class == PlmClass.PART or plm_class == PlmClass.PART_PURCH: return "(PRT)"
        if plm_class == PlmClass.DRAWING: return "(DRW)"
        return ""

    def _determine_plm_class(self, file_path: str) -> PlmClass:
        ext = os.path.splitext(file_path)[1].lower()
        path_lower = file_path.lower()
        
        if ext in [".par", ".psm"]:
            if any(lib in path_lower for lib in ["bibliothèque", "bibliotheque", "library"]):
                return PlmClass.PART_PURCH
            return PlmClass.PART
        return PlmClass.FOLDER

    def _map_to_export_row(self, cad_file: CadFile, metadata, level: int, quantity: int = 1, relationship: str = "") -> ExportRow:
        """Converts CadFile and Metadata into a flat ExportRow."""
        idx1, idx2 = revision_service.calculate_previous_indices(metadata.version)
        
        suffix = self._get_attachment_suffix(cad_file.plm_class)
        attachments = (cad_file.full_path_str + suffix)
        
        # RÈGLE : Version par défaut à "-" si vide
        version = metadata.version
        if not version.strip():
            version = "-"
            
        row = ExportRow(
            level=level,
            relationship=relationship,
            order=self._order_counter,
            quantity=quantity,
            repere="",
            special_cad=cad_file.name_without_extension,
            plm_class=cad_file.plm_class.value,
            ref_utilisat=cad_file.name_without_extension,
            version=version,
            indice_1=idx1,
            indice_2=idx2,
            revision=metadata.revision,
            designation=metadata.designation,
            cus_createur=metadata.auteur,
            cus_date_crea=metadata.date_creation,
            user_version_1=metadata.auteur_modif,
            date_version_1=metadata.date_modif,
            matiere=metadata.matiere,
            densite=metadata.densite,
            dia_se=metadata.dia_se,
            mode_appro=assembly_service.get_mode_appro(cad_file.full_path_str),
            attachments=attachments
        )
        self._order_counter += 1
        return row

    def execute(
        self,
        input_dir: str,
        output_dir: str,
        output_name: str,
        recursive: bool = False,
        dft_folder: Optional[str] = None,
        search_mode: str = "les_deux",
        progress_callback: Optional[Callable[[int, int, str], None]] = None,
        is_cancelled: Optional[Callable[[], bool]] = None
    ):
        self._order_counter = 1
        logger.info(f"Début de l'extraction par lot pour : {input_dir}")

        if progress_callback: progress_callback(5, 100, "Scan du dossier...")

        # 1. Collect PAR/PSM files
        target_files = []
        try:
            if recursive:
                for root, _, files in os.walk(input_dir):
                    for f in files:
                        if f.lower().endswith(('.par', '.psm')):
                            target_files.append(os.path.join(root, f))
            else:
                with os.scandir(input_dir) as it:
                    for entry in it:
                        if entry.is_file() and entry.name.lower().endswith(('.par', '.psm')):
                            target_files.append(entry.path)
        except Exception as e:
            logger.error(f"Erreur lors du scan du dossier : {e}")
            raise

        if not target_files:
            logger.warning("Aucun fichier .par ou .psm trouvé dans le dossier.")
            if progress_callback: progress_callback(100, 100, "Terminé (Aucun fichier)")
            return

        if is_cancelled and is_cancelled(): return

        # 2. Index Drawings
        if progress_callback: progress_callback(15, 100, "Indexation des plans...")
        index_plans = dft_indexer.index_drawings(
            seed_paths=target_files, 
            specific_folder=dft_folder,
            mode=search_mode,
            callback_progress=lambda s, f, t: progress_callback(min(15 + (s // 100), 30), 100, f"Recherche... {f} plans trouvés") if progress_callback else None
        )

        if is_cancelled and is_cancelled(): return

        # 3. Process each file and its drawing
        if progress_callback: progress_callback(35, 100, "Extraction des métadonnées...")
        
        export_rows: List[ExportRow] = []
        total = len(target_files)
        
        for idx, file_path in enumerate(target_files):
            if is_cancelled and is_cancelled(): break
            
            if idx % 10 == 0:
                pct = 35 + int((idx / total) * 50)
                progress_callback(pct, 100, f"Traitement {idx+1}/{total} : {os.path.basename(file_path)}")
            
            name_no_ext = os.path.splitext(os.path.basename(file_path))[0]
            name_lower = name_no_ext.lower()
            
            plm_class = self._determine_plm_class(file_path)
            cad_file = CadFile(file_path=Path(file_path), plm_class=plm_class)
            meta_piece = metadata_service.get_metadata_from_file(file_path)
            
            if name_lower in index_plans:
                # RULE: 0 DFT, 1 Piece
                dft_path = index_plans[name_lower]
                dft_cad = CadFile(file_path=Path(dft_path), plm_class=PlmClass.DRAWING)
                dft_meta = metadata_service.get_metadata_from_file(dft_path)
                
                # Fallbacks for DFT
                auteur = dft_meta.auteur if dft_meta.auteur.strip() else meta_piece.auteur
                date_crea = dft_meta.date_creation if dft_meta.date_creation.strip() else meta_piece.date_creation
                
                # RÈGLE : Version par défaut à "-" si vide
                version = dft_meta.version
                if not version.strip() or version == "-":
                    version = meta_piece.version
                
                if not version.strip():
                    version = "-"
                    
                # Level 0: DFT
                dft_row = ExportRow(
                    level=0,
                    relationship="",
                    order=self._order_counter,
                    quantity=1,
                    repere="",
                    special_cad=dft_cad.name_without_extension,
                    plm_class=PlmClass.DRAWING.value,
                    ref_utilisat=dft_cad.name_without_extension,
                    version=version,
                    indice_1="-",
                    indice_2="-",
                    revision=dft_meta.revision,
                    designation=meta_piece.designation,
                    cus_createur=auteur,
                    cus_date_crea=date_crea,
                    user_version_1=dft_meta.auteur_modif,
                    date_version_1=dft_meta.date_modif,
                    matiere=meta_piece.matiere,
                    densite=meta_piece.densite,
                    dia_se=meta_piece.dia_se,
                    mode_appro=assembly_service.get_mode_appro(dft_path),
                    attachments=(dft_cad.full_path_str + "(DRW)")
                )
                self._order_counter += 1
                export_rows.append(dft_row)
                
                # Level 1: Piece
                piece_row = self._map_to_export_row(cad_file, meta_piece, level=1, relationship=RelationshipType.DRAWING.value)
                export_rows.append(piece_row)
            else:
                # RULE: 0 Piece only
                piece_row = self._map_to_export_row(cad_file, meta_piece, level=0)
                export_rows.append(piece_row)

        if is_cancelled and is_cancelled(): return

        # 4. Export to Excel
        if progress_callback: progress_callback(90, 100, "Génération du fichier Excel...")
        full_output_path = os.path.join(output_dir, output_name)
        excel_exporter.create_export(export_rows, full_output_path)

        if progress_callback: progress_callback(100, 100, "Extraction terminée !")
        logger.success(f"Export par lot réussi : {output_name}")

# Global instance
batch_export_use_case = BatchExportUseCase()
