import os
from typing import List
from openpyxl import Workbook
from openpyxl.styles import PatternFill
from SE_to_PLM.core.models.export_row import ExportRow
from SE_to_PLM.core.services.export.excel_styles import apply_header_style, apply_data_style
from SE_to_PLM.core.services.export.column_auto_size import auto_size_columns
from SE_to_PLM.core.services.plm.abbreviation_service import abbreviation_service
from SE_to_PLM.core.services.logging.logger_service import logger

class ExcelExporter:
    """
    Service for generating styled Excel files from ExportRow data.
    """
    
    HEADERS = [
        "Level", "Relationship", "ordre", "quantite", "repere", "SpecialCAD",
        "Class", "ref_utilisat", "version", "indice_1", "indice_2", "revision",
        "designation", "designation_erp", "cus_createur", "cus_date_crea", "user_version_1",
        "date_version_1", "matiere", "densite", "dia_se", "mode_appro", "Attachments"
    ]

    def create_export(self, rows: List[ExportRow], output_path: str,
                      optimize_designations: bool = False,
                      highlight_modifications: bool = False):
        """
        Creates a new Excel workbook, writes the data, applies styles, 
        optimizes designations with abbreviations if enabled to a separate
        designation_erp column, and saves it.
        """
        logger.info(f"Création du fichier Excel...")
        
        # Ensure directory exists
        os.makedirs(os.path.dirname(output_path), exist_ok=True)
        
        wb = Workbook()
        ws = wb.active
        ws.title = "Structure"

        # 1. Write Headers
        for col_idx, header in enumerate(self.HEADERS, 1):
            cell = ws.cell(row=1, column=col_idx, value=header)
            apply_header_style(cell, col_idx)

        # 2. Write Data Rows
        fill_modified = PatternFill(fill_type="solid", fgColor="FF9999")
        
        for row_idx, row_obj in enumerate(rows, 2):
            row_data = row_obj.to_list()
            row_was_modified = False
            
            orig_designation = row_data[12] # index 12 corresponds to designation field
            
            erp_designation = orig_designation
            if optimize_designations and isinstance(orig_designation, str):
                optimized, modified = abbreviation_service.optimiser_designation(orig_designation)
                erp_designation = abbreviation_service.majuscules_sans_accents(optimized)
                if erp_designation != orig_designation:
                    row_was_modified = True
                    
            # Insert designation_erp right after designation
            new_row_data = row_data[:13] + [erp_designation] + row_data[13:]
            
            for col_idx, value in enumerate(new_row_data, 1):
                cell = ws.cell(row=row_idx, column=col_idx, value=value)
                apply_data_style(cell)
                
            if row_was_modified and highlight_modifications:
                for c_idx in range(1, len(self.HEADERS) + 1):
                    ws.cell(row=row_idx, column=c_idx).fill = fill_modified

        # 3. Auto-size columns
        try:
            auto_size_columns(ws)
        except Exception:
            pass

        # 4. Save
        try:
            # Ensure .xlsx extension
            if not output_path.lower().endswith(".xlsx"):
                output_path += ".xlsx"
                
            wb.save(output_path)
        except Exception as e:
            logger.error(f"Erreur d'enregistrement Excel : {e}")
            raise

# Global instance
excel_exporter = ExcelExporter()
