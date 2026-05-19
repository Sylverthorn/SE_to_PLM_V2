import os
from typing import List, Dict, Optional, Callable, Set, Tuple
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

class UnifiedMultiASMExportUseCase:
    """
    Consolidated use case for Multi-ASM export.
    Supports:
    - Source: Single ASM file OR Directory of ASM files.
    - Output: Multiple Excel files (one per ASM) OR Single Excel file (Blocks of Niv 0/1).
    """

    def __init__(self):
        self._order_counter = 1

    def _get_attachment_suffix(self, plm_class: PlmClass) -> str:
        if plm_class == PlmClass.SUB_ASSY: return "(ASM)"
        if plm_class == PlmClass.PART or plm_class == PlmClass.PART_PURCH: return "(PRT)"
        if plm_class == PlmClass.DRAWING: return "(DRW)"
        return ""

    def _map_to_export_row(self, node: AssemblyNode, level: int, relationship: str = "") -> ExportRow:
        """Converts an AssemblyNode into a flat ExportRow."""
        idx1, idx2 = revision_service.calculate_previous_indices(node.metadata.version)
        
        suffix = self._get_attachment_suffix(node.cad_file.plm_class)
        attachments = (node.cad_file.full_path_str + suffix)
        
        # RÈGLE : Version par défaut à "-" si vide
        version = node.metadata.version
        if not version.strip():
            version = "-"
            
        row = ExportRow(
            level=level,
            relationship=relationship,
            order=self._order_counter,
            quantity=node.quantity if level > 0 else 1,
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
            attachments=attachments
        )
        self._order_counter += 1
        return row

    def execute(
        self,
        input_path: str,
        output_dir: str,
        output_name: str,
        is_folder_source: bool = False,
        output_mode: str = "single", # "single" or "multiple"
        dft_folder: Optional[str] = None,
        search_mode: str = "les_deux",
        progress_callback: Optional[Callable[[int, int, str], None]] = None,
        is_cancelled: Optional[Callable[[], bool]] = None
    ):
        self._order_counter = 1
        source_type = "dossier" if is_folder_source else "fichier"
        logger.info(f"Début Multi-ASM (Source: {source_type}, Format: {output_mode})")

        # 1. Collect Head ASM files
        head_asms = []
        if is_folder_source:
            if progress_callback: progress_callback(2, 100, "Scan du dossier source...")
            for root, _, files in os.walk(input_path):
                for f in files:
                    if f.lower().endswith(".asm"):
                        head_asms.append(os.path.join(root, f))
        else:
            head_asms.append(input_path)

        if not head_asms:
            logger.warning("Aucun fichier .asm trouvé.")
            return

        # 2. Index Drawings (Important: always index for Multi-ASM)
        if progress_callback: progress_callback(5, 100, "Recherche des plans...")
        index_plans = dft_indexer.index_drawings(
            seed_paths=head_asms, 
            specific_folder=dft_folder or (input_path if is_folder_source else None), 
            mode=search_mode,
            callback_progress=lambda s, f, t: progress_callback(min(5 + (s // 100), 15), 100, f"Recherche... {f} plans trouvés") if progress_callback else None
        )

        if is_cancelled and is_cancelled(): return

        # 3. Connect to Solid Edge
        if progress_callback: progress_callback(20, 100, "Connexion à Solid Edge...")
        app = connection_manager.get_application()
        if not app:
            logger.error("Impossible de se connecter à Solid Edge.")
            return

        # 4. Explore all assemblies (Deduplicated)
        unique_assemblies: Dict[str, AssemblyNode] = {}
        unique_items_for_dft: Dict[str, AssemblyNode] = {}
        
        total_heads = len(head_asms)
        for idx, head_path in enumerate(head_asms):
            if is_cancelled and is_cancelled(): break
            
            progress_val = 25 + int((idx / total_heads) * 30)
            if progress_callback: progress_callback(progress_val, 100, f"Analyse structure {idx+1}/{total_heads}...")
            
            try:
                doc = app.Documents.Open(head_path)
                tree = assembly_service.explore_assembly(doc)
                
                def collect_data(node: AssemblyNode):
                    path_lower = node.cad_file.full_path_str.lower()
                    unique_items_for_dft[path_lower] = node
                    
                    if node.cad_file.plm_class == PlmClass.SUB_ASSY:
                        if path_lower not in unique_assemblies:
                            unique_assemblies[path_lower] = node
                    
                    for child in node.children:
                        collect_data(child)

                collect_data(tree)
                doc.Close(False) # Always close after analysis
            except Exception as e:
                logger.warning(f"Impossible d'analyser {os.path.basename(head_path)} : {e}")

        if is_cancelled and is_cancelled(): return

        # 5. Generation
        if output_mode == "multiple":
            # OUTPUT: Folder with multiple Excel files
            full_export_path = os.path.join(output_dir, output_name)
            os.makedirs(full_export_path, exist_ok=True)
            
            total_asm = len(unique_assemblies)
            for idx, (path, asm_node) in enumerate(unique_assemblies.items()):
                if is_cancelled and is_cancelled(): break
                
                pct = 60 + int((idx / total_asm) * 40)
                if progress_callback: progress_callback(pct, 100, f"Génération {idx+1}/{total_asm}...")
                
                rows: List[ExportRow] = []
                self._order_counter = 1
                
                # Assembly Rows
                rows.append(self._map_to_export_row(asm_node, level=0))
                for child in asm_node.children:
                    rows.append(self._map_to_export_row(child, level=1, relationship="ComposedOf"))
                
                # Add Drawings for THIS assembly and its children
                processed_plans = set()
                # 1. Drawing for the assembly itself
                self._check_and_add_drawing(asm_node, index_plans, processed_plans, rows)
                # 2. Drawings for children
                for child in asm_node.children:
                    self._check_and_add_drawing(child, index_plans, processed_plans, rows)
                
                excel_path = os.path.join(full_export_path, f"{asm_node.cad_file.name_without_extension}.xlsx")
                excel_exporter.create_export(rows, excel_path)
            
            logger.success(f"Export Multi-ASM (Fichiers multiples) terminé dans : {output_name}")

        else:
            # OUTPUT: Single Excel file with blocks
            export_rows: List[ExportRow] = []
            self._order_counter = 1
            
            total_asm = len(unique_assemblies)
            for idx, (path, asm_node) in enumerate(unique_assemblies.items()):
                if is_cancelled and is_cancelled(): break
                
                pct = 60 + int((idx / total_asm) * 25)
                if progress_callback: progress_callback(pct, 100, f"Bloc ASM {idx+1}/{total_asm}...")
                
                export_rows.append(self._map_to_export_row(asm_node, level=0))
                for child in asm_node.children:
                    export_rows.append(self._map_to_export_row(child, level=1, relationship="ComposedOf"))

            # Plans at the end
            if progress_callback: progress_callback(85, 100, "Traitement des plans...")
            processed_plans = set()
            for path_lower, node in unique_items_for_dft.items():
                if is_cancelled and is_cancelled(): break
                self._check_and_add_drawing(node, index_plans, processed_plans, export_rows)

            if progress_callback: progress_callback(95, 100, "Sauvegarde du fichier...")
            full_output_path = os.path.join(output_dir, output_name)
            if not full_output_path.lower().endswith(".xlsx"): full_output_path += ".xlsx"
            excel_exporter.create_export(export_rows, full_output_path)
            
            logger.success(f"Export Multi-ASM (Fichier unique) terminé : {output_name}")

        if progress_callback: progress_callback(100, 100, "Extraction terminée !")

    def _check_and_add_drawing(self, node: AssemblyNode, index_plans: Dict[str, str], processed_plans: Set[str], rows: List[ExportRow]):
        name_lower = node.cad_file.name_without_extension.lower()
        if name_lower in index_plans:
            dft_path = index_plans[name_lower]
            if dft_path not in processed_plans:
                processed_plans.add(dft_path)
                self._add_drawing_rows(dft_path, node, rows)

    def _add_drawing_rows(self, dft_path: str, source_node: AssemblyNode, rows: List[ExportRow]):
        """Standard 2-row block for a drawing."""
        dft_cad = CadFile(file_path=Path(dft_path), plm_class=PlmClass.DRAWING)
        dft_meta = metadata_service.get_metadata_from_file(dft_path)
        
        auteur = dft_meta.auteur if dft_meta.auteur.strip() else source_node.metadata.auteur
        date_crea = dft_meta.date_creation if dft_meta.date_creation.strip() else source_node.metadata.date_creation
        
        # RÈGLE : Version par défaut à "-" si vide
        version = dft_meta.version
        if not version.strip() or version == "-":
            version = source_node.metadata.version
            
        if not version.strip():
            version = "-"
            
        dft_row = ExportRow(
            level=0, relationship="", order=self._order_counter, quantity=1, repere="",
            special_cad=dft_cad.name_without_extension,
            plm_class=PlmClass.DRAWING.value,
            ref_utilisat=dft_cad.name_without_extension,
            version=version,
            indice_1="-", indice_2="-", revision=dft_meta.revision,
            designation=source_node.metadata.designation,
            cus_createur=auteur, cus_date_crea=date_crea,
            user_version_1=dft_meta.auteur_modif, date_version_1=dft_meta.date_modif,
            matiere=source_node.metadata.matiere, densite=source_node.metadata.densite,
            dia_se=source_node.metadata.dia_se,
            attachments=(dft_cad.full_path_str + "(DRW)")
        )
        self._order_counter += 1
        rows.append(dft_row)
        
        # Piece Row
        piece_row = self._map_to_export_row(source_node, level=1, relationship=RelationshipType.DRAWING.value)
        rows.append(piece_row)

# Global instance
unified_multi_asm_export_use_case = UnifiedMultiASMExportUseCase()
