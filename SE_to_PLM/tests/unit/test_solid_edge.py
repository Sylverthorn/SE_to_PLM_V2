import pytest
from SE_to_PLM.core.services.solid_edge.property_reader import PropertyReader
from SE_to_PLM.core.services.solid_edge.assembly_service import AssemblyService
from SE_to_PLM.core.models.enums import PlmClass

def test_date_normalization():
    reader = PropertyReader()
    assert reader.normalize_date("14/05/2026") == "14/05/2026 12:00:00 AM"
    assert reader.normalize_date("2026-05-14") == "14/05/2026 12:00:00 AM"
    assert reader.normalize_date("14.05.2026") == "14/05/2026 12:00:00 AM"
    assert reader.normalize_date("invalid") == "invalid"

def test_plm_class_determination():
    service = AssemblyService()
    assert service.determine_plm_class("test.asm", is_root=True) == PlmClass.SUB_ASSY
    assert service.determine_plm_class("sub.asm") == PlmClass.SUB_ASSY
    assert service.determine_plm_class("part.par") == PlmClass.PART
    assert service.determine_plm_class("sheet.psm") == PlmClass.PART
    assert service.determine_plm_class("C:/Library/bolt.par") == PlmClass.PART_PURCH
    assert service.determine_plm_class("drawing.dft") == PlmClass.DRAWING
    assert service.determine_plm_class("unknown.txt") == PlmClass.FOLDER
