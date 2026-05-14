import os
from pathlib import Path

# Chemins de base
APP_ROOT = Path(__file__).parent.parent
DEFAULT_EXPORT_DIR = Path.home() / "Documents" / "Exports_PLM"

# Extensions supportées
ASM_EXT = ".asm"
PAR_EXT = ".par"
PSM_EXT = ".psm"
DFT_EXT = ".dft"

# Classes PLM
CLASS_SUB_ASSY = "SUB_ASSY_A"
CLASS_PART = "PART_A"
CLASS_PART_PURCH = "PART_PURCH_A"
CLASS_DRAWING = "CAD_DRAWING_A"
CLASS_FOLDER = "Folder"

# Configuration Cache
METADATA_CACHE_SIZE = 1000

# Modes de recherche DFT
SEARCH_MODE_ARBO = "arborescence"
SEARCH_MODE_SPECIFIC = "dossier_specifique"
SEARCH_MODE_BOTH = "les_deux"
