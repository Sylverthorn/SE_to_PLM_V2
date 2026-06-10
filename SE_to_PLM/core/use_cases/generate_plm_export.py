import os
import win32com.client
from typing import List, Dict, Optional, Callable
from pathlib import Path

from SE_to_PLM.core.models.enums import PlmClass, RelationshipType
from SE_to_PLM.core.models.cad_file import CadFile
from SE_to_PLM.core.models.assembly_node import AssemblyNode
from SE_to_PLM.core.models.export_row import ExportRow
from SE_to_PLM.infrastructure.solid_edge.connection_manager import connection_manager
from SE_to_PLM.core.services.solid_edge.assembly_service import assembly_service
from SE_to_PLM.core.services.solid_edge.metadata_service import metadata_service
from SE_to_PLM.core.services.indexing.dft_indexer import dft_indexer
from SE_to_PLM.core.services.export.excel_exporter import excel_exporter
from SE_to_PLM.core.services.plm.revision_service import revision_service
from SE_to_PLM.core.services.logging.logger_service import logger

class GeneratePLMExportUseCase:
    """
    Main orchestrator for the PLM export workflow.
    """

    def __init__(self):
        self._order_counter = 1

    def _get_attachment_suffix(self, plm_class: PlmClass) -> str:
        if plm_class == PlmClass.SUB_ASSY: return "(ASM)"
        if plm_class == PlmClass.PART or plm_class == PlmClass.PART_PURCH: return "(PRT)"
        if plm_class == PlmClass.DRAWING: return "(DRW)"
        return ""

    def _map_to_export_row(self, node: AssemblyNode) -> ExportRow:
        """Converts an AssemblyNode into a flat ExportRow."""
        idx1, idx2 = revision_service.calculate_previous_indices(node.metadata.version)
        
        suffix = self._get_attachment_suffix(node.cad_file.plm_class)
        attachments = (node.cad_file.full_path_str + suffix)
        
        # RÈGLE : Version par défaut à "-" si vide
        version = node.metadata.version
        if not version.strip():
            version = "-"
            
        row = ExportRow(
            level=node.level,
            relationship=node.relationship.value,
            order=self._order_counter,
            quantity=node.quantity,
            repere="",
            special_cad=node.cad_file.name_without_extension,
            plm_class=node.cad_file.plm_class.value,
            ref_utilisat=node.cad_file.name_without_extension,
            version=version,
            indice_1=idx1,
            indice_2=idx2,
            revision=node.metadata.revision,
            designation=node.metadata.designation,
            cus_createur=node.metadata.auteur,
            cus_date_crea=node.metadata.date_creation,
            user_version_1=node.metadata.auteur_modif,
            date_version_1=node.metadata.date_modif,
            matiere=node.metadata.matiere,
            densite=node.metadata.densite,
            dia_se=node.metadata.dia_se,
            mode_appro=assembly_service.get_mode_appro(node.cad_file.full_path_str),
            attachments=attachments
        )
        self._order_counter += 1
        return row

    def execute(
        self,
        input_file: str,
        output_dir: str,
        output_name: str,
        dft_folder: Optional[str] = None,
        search_mode: str = "les_deux",
        progress_callback: Optional[Callable[[int, int, str], None]] = None,
        is_cancelled: Optional[Callable[[], bool]] = None
    ):
        self._order_counter = 1
        logger.info(f"Début de l'extraction pour : {os.path.basename(input_file)}")

        if progress_callback: progress_callback(5, 100, "Recherche des plans...")

        # 1. Index Drawings
        index_plans = dft_indexer.index_drawings(
            seed_paths=[input_file], 
            specific_folder=dft_folder, 
            mode=search_mode,
            callback_progress=lambda s, f, t: progress_callback(min(5 + (s // 100), 20), 100, f"Recherche... {f} plans trouvés") if progress_callback else None
        )

        if is_cancelled and is_cancelled(): return

        # 2. Connect to Design Manager (Revision Manager) instead of Solid Edge
        if progress_callback: progress_callback(25, 100, "Démarrage du gestionnaire de liens...")
        try:
            rm_app = win32com.client.Dispatch("RevisionManager.Application")
            rm_app.Visible = False
        except Exception as e:
            logger.error(f"Impossible de lancer Revision Manager : {e}")
            return

        # 3. Traverse Assembly
        if progress_callback: progress_callback(35, 100, "Analyse de la structure...")
        try:
            tree = assembly_service.explore_assembly(input_file, rm_app)
            
            if is_cancelled and is_cancelled(): 
                try: rm_app.Quit()
                except: pass
                return

            # 4. Flatten and Process
            if progress_callback: progress_callback(70, 100, "Préparation des données...")
            
            assembly_rows: List[ExportRow] = []
            drawing_rows: List[ExportRow] = []
            processed_plans = set()
            
            def process_node(node: AssemblyNode):
                assembly_rows.append(self._map_to_export_row(node))
                
                name_lower = node.cad_file.name_without_extension.lower()
                if name_lower in index_plans:
                    dft_path = index_plans[name_lower]
                    if dft_path not in processed_plans:
                        processed_plans.add(dft_path)
                        pending_plans.append((dft_path, node))
                
                for child in node.children:
                    process_node(child)

            pending_plans = []
            process_node(tree)
            
            for dft_path, source_node in pending_plans:
                self._add_drawing_rows(dft_path, source_node, drawing_rows)

            export_rows = assembly_rows + drawing_rows

            # 5. Export to Excel
            if progress_callback: progress_callback(90, 100, "Génération du fichier Excel...")
            full_output_path = os.path.join(output_dir, output_name)
            excel_exporter.create_export(export_rows, full_output_path)

            if progress_callback: progress_callback(100, 100, "Extraction terminée !")
            logger.success(f"Export réussi : {output_name}")

        except Exception as e:
            logger.error(f"Échec de l'extraction : {str(e)}")
            raise
        finally:
            try:
                rm_app.Quit()
            except:
                pass

    def _add_drawing_rows(self, dft_path: str, source_node: AssemblyNode, rows: List[ExportRow]):
        """
        Adds the special 2-row entry for a drawing:
        1. The DFT itself (Level 0)
        2. The link to the source piece (Level 1, Relationship='Drawing')
        """
        # DFT Row (Level 0)
        dft_cad = CadFile(file_path=Path(dft_path), plm_class=PlmClass.DRAWING)
        
        # Extract actual DFT metadata
        dft_meta = metadata_service.get_metadata_from_file(dft_path)
        
        # Fallback to source piece for empty critical fields as per legacy behavior
        auteur = dft_meta.auteur if dft_meta.auteur.strip() else source_node.metadata.auteur
        date_crea = dft_meta.date_creation if dft_meta.date_creation.strip() else source_node.metadata.date_creation
        
        # RÈGLE : Version par défaut à "-" si vide
        version = dft_meta.version
        if not version.strip() or version == "-":
            version = source_node.metadata.version
            
        if not version.strip():
            version = "-"
            
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
            designation=source_node.metadata.designation, # Usually same as piece
            cus_createur=auteur,
            cus_date_crea=date_crea,
            user_version_1=dft_meta.auteur_modif,
            date_version_1=dft_meta.date_modif,
            matiere=source_node.metadata.matiere, # DFTs don't have material
            densite=source_node.metadata.densite,
            dia_se=source_node.metadata.dia_se,
            mode_appro=assembly_service.get_mode_appro(source_node.cad_file.full_path_str),
            attachments=(dft_cad.full_path_str + "(DRW)")
        )
        self._order_counter += 1
        rows.append(dft_row)
        
        # Reference Row (Level 1)
        ref_row = self._map_to_export_row(source_node)
        ref_row.level = 1
        ref_row.relationship = RelationshipType.DRAWING.value
        rows.append(ref_row)

# Global instance
generate_plm_export_use_case = GeneratePLMExportUseCase()
