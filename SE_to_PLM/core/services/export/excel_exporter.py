import os
from typing import List
from openpyxl import Workbook
from SE_to_PLM.core.models.export_row import ExportRow
from SE_to_PLM.core.services.export.excel_styles import apply_header_style, apply_data_style
from SE_to_PLM.core.services.export.column_auto_size import auto_size_columns
from SE_to_PLM.core.services.logging.logger_service import logger

class ExcelExporter:
    """
    Service for generating styled Excel files from ExportRow data.
    """
    
    HEADERS = [
        "Level", "Relationship", "ordre", "quantite", "repere", "SpecialCAD",
        "Class", "ref_utilisat", "version", "indice_1", "indice_2", "revision",
        "designation", "cus_createur", "cus_date_crea", "user_version_1",
        "date_version_1", "matiere", "densite", "dia_se", "Attachments"
    ]

    def create_export(self, rows: List[ExportRow], output_path: str):
        """
        Creates a new Excel workbook, writes the data, applies styles, 
        and saves it to the specified path.
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
        for row_idx, row_obj in enumerate(rows, 2):
            row_data = row_obj.to_list()
            for col_idx, value in enumerate(row_data, 1):
                cell = ws.cell(row=row_idx, column=col_idx, value=value)
                apply_data_style(cell)

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
