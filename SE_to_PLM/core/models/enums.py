from enum import Enum

class PlmClass(Enum):
    SUB_ASSY = "SUB_ASSY_A"
    PART = "PART_A"
    PART_PURCH = "PART_PURCH_A"
    DRAWING = "CAD_DRAWING_A"
    FOLDER = "Folder"

class RelationshipType(Enum):
    NONE = ""
    COMPOSED_OF = "ComposedOf"
    DRAWING = "Drawing"

class CadExtension(Enum):
    ASM = ".asm"
    PAR = ".par"
    PSM = ".psm"
    DFT = ".dft"
