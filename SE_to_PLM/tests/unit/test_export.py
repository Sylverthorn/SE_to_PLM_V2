import os
import pytest
from SE_to_PLM.core.models.export_row import ExportRow
from SE_to_PLM.core.services.export.excel_exporter import ExcelExporter
from openpyxl import load_workbook

def test_excel_generation(tmp_path):
    output_file = str(tmp_path / "test_export.xlsx")
    
    rows = [
        ExportRow(
            level=0, relationship="", order=1, quantity=1, repere="",
            special_cad="Root", plm_class="SUB_ASSY_A", ref_utilisat="Root",
            version="A", indice_1="-", indice_2="-", revision="1",
            designation="Top Assy", cus_createur="Admin", cus_date_crea="2026",
            user_version_1="", date_version_1="", matiere="",
            densite="", dia_se="", mode_appro="fabrication interne", attachments="C:/Root.asm"
        ),
        ExportRow(
            level=1, relationship="ComposedOf", order=2, quantity=2, repere="",
            special_cad="Part1", plm_class="PART_A", ref_utilisat="Part1",
            version="B", indice_1="A", indice_2="-", revision="1",
            designation="Simple Part", cus_createur="User", cus_date_crea="2026",
            user_version_1="", date_version_1="", matiere="Steel",
            densite="7.8", dia_se="", mode_appro="fabrication interne", attachments="C:/Part1.par"
        )
    ]
    
    exporter = ExcelExporter()
    exporter.create_export(rows, output_file)
    
    assert os.path.exists(output_file)
    
    # Verify content
    wb = load_workbook(output_file)
    ws = wb.active
    assert ws.title == "Structure"
    assert ws.cell(row=1, column=1).value == "Level"
    assert ws.cell(row=2, column=6).value == "Root"
    assert ws.cell(row=3, column=6).value == "Part1"
    assert ws.cell(row=3, column=4).value == 2 # quantity
    
    # Verify some styling (limited check)
    assert ws.cell(row=1, column=1).font.bold is True
    # Column 1 (Green) vs Column 13 (Orange)
    assert ws.cell(row=1, column=1).fill.start_color.index == "00CCFFCC"
    assert ws.cell(row=1, column=13).fill.start_color.index == "00FFC000"
