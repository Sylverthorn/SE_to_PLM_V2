import pytest
from pathlib import Path
from SE_to_PLM.core.models.enums import PlmClass, RelationshipType
from SE_to_PLM.core.models.metadata import Metadata
from SE_to_PLM.core.models.cad_file import CadFile
from SE_to_PLM.core.models.assembly_node import AssemblyNode
from SE_to_PLM.core.models.export_row import ExportRow

def test_metadata_creation():
    meta = Metadata(designation="Test Part", revision="2")
    assert meta.designation == "Test Part"
    assert meta.revision == "2"
    assert meta.version == "-"

def test_cad_file_properties():
    path = Path("C:/Projects/Part1.par")
    cad = CadFile(file_path=path, plm_class=PlmClass.PART)
    assert cad.name_without_extension == "Part1"
    assert cad.extension == ".par"
    assert "Part1.par" in cad.full_path_str

def test_assembly_node_hierarchy():
    root_cad = CadFile(Path("Root.asm"), PlmClass.SUB_ASSY)
    child_cad = CadFile(Path("Child.par"), PlmClass.PART)
    
    root_node = AssemblyNode(cad_file=root_cad, metadata=Metadata.default())
    child_node = AssemblyNode(
        cad_file=child_cad, 
        metadata=Metadata.default(), 
        level=1, 
        relationship=RelationshipType.COMPOSED_OF
    )
    
    root_node.add_child(child_node)
    
    assert len(root_node.children) == 1
    assert root_node.children[0].cad_file.name_without_extension == "Child"

def test_export_row_to_list():
    row = ExportRow(
        level=0, relationship="", order=1, quantity=1, repere="",
        special_cad="Test", plm_class="PART_A", ref_utilisat="Test",
        version="A", indice_1="-", indice_2="-", revision="1",
        designation="Desc", cus_createur="User", cus_date_crea="Date",
        user_version_1="", date_version_1="", matiere="Steel",
        densite="7.8", dia_se="", attachments="Path"
    )
    row_list = row.to_list()
    assert len(row_list) == 21
    assert row_list[0] == 0
    assert row_list[6] == "PART_A"
    assert row_list[20] == "Path"
