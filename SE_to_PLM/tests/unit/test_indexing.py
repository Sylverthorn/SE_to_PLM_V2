import os
import shutil
import tempfile
import pytest
from SE_to_PLM.core.services.indexing.dft_indexer import DFTIndexer
from SE_to_PLM.core.services.indexing.path_resolver import PathResolver

@pytest.fixture
def temp_project_structure():
    # Create a dummy project structure
    # root/
    #   L1/
    #     L2/
    #       L3/
    #         main.asm
    #   drawings/
    #     part1.dft
    #     sub1.dft
    temp_dir = tempfile.mkdtemp()
    l3_dir = os.path.join(temp_dir, "L1", "L2", "L3")
    drw_dir = os.path.join(temp_dir, "drawings")
    
    os.makedirs(l3_dir)
    os.makedirs(drw_dir)
    
    asm_path = os.path.join(l3_dir, "main.asm")
    open(asm_path, 'w').close()
    
    open(os.path.join(drw_dir, "part1.dft"), 'w').close()
    open(os.path.join(drw_dir, "sub1.dft"), 'w').close()
    
    yield asm_path, drw_dir
    
    shutil.rmtree(temp_dir)

def test_path_resolver_root(temp_project_structure):
    asm_path, _ = temp_project_structure
    resolver = PathResolver()
    root = resolver.get_project_root(asm_path, levels_up=3)
    
    # root should be the temp_dir
    assert os.path.basename(root).startswith("tmp") or "tmp" in root.lower()
    assert os.path.exists(os.path.join(root, "drawings"))

def test_indexer_scan(temp_project_structure):
    asm_path, drw_dir = temp_project_structure
    indexer = DFTIndexer()
    
    # Scan with Both mode
    index = indexer.index_drawings(asm_path, specific_folder=drw_dir, mode="les_deux")
    
    assert "part1" in index
    assert "sub1" in index
    assert index["part1"].lower().endswith("part1.dft")
    assert len(index) == 2

def test_path_resolver_network_normalization():
    resolver = PathResolver()
    path = "\\\\192.168.1.10\\Shared\\Project/Sub"
    norm = resolver.normalize_network_path(path)
    assert norm == "\\\\192.168.1.10\\Shared\\Project\\Sub"
