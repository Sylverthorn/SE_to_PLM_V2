import os
from pathlib import Path
from typing import List, Dict, Any, Optional
from SE_to_PLM.core.models.enums import PlmClass, RelationshipType
from SE_to_PLM.core.models.cad_file import CadFile
from SE_to_PLM.core.models.assembly_node import AssemblyNode
from SE_to_PLM.core.services.solid_edge.metadata_service import metadata_service
from SE_to_PLM.core.services.logging.logger_service import logger

class AssemblyService:
    """
    Handles assembly structure traversal and PLM class determination.
    """

    def determine_plm_class(self, file_path: str, is_root: bool = False) -> PlmClass:
        """
        Determines the PLM class based on the file extension and path.
        """
        ext = os.path.splitext(file_path)[1].lower()
        path_lower = file_path.lower()

        if is_root or ext == ".asm":
            return PlmClass.SUB_ASSY
        
        if ext in [".par", ".psm"]:
            if any(lib in path_lower for lib in ["bibliothèque", "bibliotheque", "library"]):
                return PlmClass.PART_PURCH
            return PlmClass.PART
        
        if ext == ".dft":
            return PlmClass.DRAWING
            
        return PlmClass.FOLDER

    def explore_assembly(self, assembly_doc: Any, level: int = 0) -> AssemblyNode:
        """
        Recursively explores a Solid Edge assembly document and returns the tree structure.
        """
        file_path = assembly_doc.FullName
        plm_class = self.determine_plm_class(file_path, is_root=(level == 0))
        
        cad_file = CadFile(file_path=Path(file_path), plm_class=plm_class)
        metadata = metadata_service.get_metadata_from_object(assembly_doc)
        
        root_node = AssemblyNode(
            cad_file=cad_file,
            metadata=metadata,
            level=level,
            relationship=RelationshipType.NONE if level == 0 else RelationshipType.COMPOSED_OF
        )

        # Explore children if it's an assembly OR if it's the root node (to be safe)
        if plm_class == PlmClass.SUB_ASSY or level == 0:
            try:
                # Check if the document actually has Occurrences (only ASM does)
                if hasattr(assembly_doc, "Occurrences"):
                    occurrences = assembly_doc.Occurrences
                # Deduplicate by file path
                unique_occs: Dict[str, Dict[str, Any]] = {}
                
                for i in range(1, occurrences.Count + 1):
                    occ = occurrences.Item(i)
                    try:
                        # Try multiple ways to get the full path
                        occ_path = ""
                        try: occ_path = occ.OccurrenceDocument.FullName
                        except:
                            try: occ_path = occ.OccurrenceFileName
                            except:
                                try: occ_path = occ.FileName
                                except: pass
                        
                        if not occ_path:
                            continue
                            
                        norm_path = os.path.normcase(os.path.abspath(occ_path))
                        
                        if norm_path in unique_occs:
                            unique_occs[norm_path]["quantity"] += 1
                        else:
                            unique_occs[norm_path] = {
                                "obj": occ,
                                "quantity": 1,
                                "path": occ_path
                            }
                    except Exception as e:
                        logger.warning(f"Error accessing occurrence {i}: {e}")

                # Process unique occurrences
                for path, data in unique_occs.items():
                    occ = data["obj"]
                    occ_path = data["path"]
                    qty = data["quantity"]
                    
                    occ_plm_class = self.determine_plm_class(occ_path)
                    occ_cad = CadFile(file_path=Path(occ_path), plm_class=occ_plm_class)
                    
                    # For sub-components, we use fast extraction (FileProperties)
                    occ_metadata = metadata_service.get_metadata_from_file(occ_path)
                    
                    child_node = AssemblyNode(
                        cad_file=occ_cad,
                        metadata=occ_metadata,
                        level=level + 1,
                        quantity=qty,
                        relationship=RelationshipType.COMPOSED_OF
                    )
                    
                    # Recurse if it's a sub-assembly
                    try:
                        if occ.Subassembly:
                            sub_doc = occ.OccurrenceDocument
                            sub_tree = self.explore_assembly(sub_doc, level + 1)
                            child_node.children = sub_tree.children
                    except:
                        pass
                        
                    root_node.add_child(child_node)
                    
            except Exception as e:
                logger.error(f"Error exploring occurrences of {file_path}: {e}")

        return root_node

# Global instance
assembly_service = AssemblyService()
