from dataclasses import dataclass, field
from typing import List, Optional
from SE_to_PLM.core.models.cad_file import CadFile
from SE_to_PLM.core.models.metadata import Metadata
from SE_to_PLM.core.models.enums import RelationshipType

@dataclass
class AssemblyNode:
    cad_file: CadFile
    metadata: Metadata
    level: int = 0
    quantity: int = 1
    relationship: RelationshipType = RelationshipType.NONE
    children: List["AssemblyNode"] = field(default_factory=list)
    
    def add_child(self, child: "AssemblyNode"):
        self.children.append(child)
