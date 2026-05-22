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

def test_density_normalization():
    reader = PropertyReader()
    assert reader.normalize_density("7850,00 kg/m^3") == "7850.00"
    assert reader.normalize_density("7 850 kg/m3") == "7850"
    assert reader.normalize_density("7.85 g/cm3") == "7.85"
    assert reader.normalize_density("7850") == "7850"
    assert reader.normalize_density(None) == ""
    assert reader.normalize_density("") == ""
    assert reader.normalize_density("7850.50") == "7850.50"
    assert reader.normalize_density("1,5 g/cm3") == "1.5"
    assert reader.normalize_density("8 900,5 kg/m^3") == "8900.5"

def test_plm_class_determination():
    service = AssemblyService()
    assert service.determine_plm_class("test.asm", is_root=True) == PlmClass.SUB_ASSY
    assert service.determine_plm_class("sub.asm") == PlmClass.SUB_ASSY
    assert service.determine_plm_class("part.par") == PlmClass.PART
    assert service.determine_plm_class("sheet.psm") == PlmClass.PART
    assert service.determine_plm_class("C:/Library/bolt.par") == PlmClass.PART_PURCH
    assert service.determine_plm_class("drawing.dft") == PlmClass.DRAWING
    assert service.determine_plm_class("unknown.txt") == PlmClass.FOLDER
