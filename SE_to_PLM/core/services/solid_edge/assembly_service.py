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
    Can use either Solid Edge Application or Revision Manager.
    """

    def determine_plm_class(self, file_path: str, is_root: bool = False) -> PlmClass:
        """
        Determines the PLM class based on the file extension and path.
        """
        ext = os.path.splitext(file_path)[1].lower()
        path_lower = file_path.lower()

        if ext == ".asm":
            return PlmClass.SUB_ASSY
        
        if ext in [".par", ".psm"]:
            if any(lib in path_lower for lib in ["bibliothèque", "bibliotheque", "library"]):
                return PlmClass.PART_PURCH
            return PlmClass.PART
        
        if ext == ".dft":
            return PlmClass.DRAWING
            
        return PlmClass.FOLDER

    def explore_assembly(self, doc_or_path: Any, rm_app: Optional[Any] = None, level: int = 0) -> AssemblyNode:
        """
        Entry point for exploring an assembly structure.
        - If rm_app is provided: doc_or_path must be a file path string.
        - Otherwise: doc_or_path must be a SolidEdge.AssemblyDocument object.
        """
        if rm_app:
            return self._explore_with_revision_manager(doc_or_path, rm_app, level)
        else:
            return self._explore_with_solid_edge(doc_or_path, level)

    def _explore_with_revision_manager(self, file_path: str, rm_app: Any, level: int = 0) -> AssemblyNode:
        """
        Recursively explores assembly structure using Revision Manager (very fast).
        """
        plm_class = self.determine_plm_class(file_path, is_root=(level == 0))
        cad_file = CadFile(file_path=Path(file_path), plm_class=plm_class)
        
        # Use file properties directly for RM (safer, avoids COM Property issues on RM doc objects)
        metadata = metadata_service.get_metadata_from_file(file_path)

        root_node = AssemblyNode(
            cad_file=cad_file,
            metadata=metadata,
            level=level,
            relationship=RelationshipType.NONE if level == 0 else RelationshipType.COMPOSED_OF
        )

        if plm_class == PlmClass.SUB_ASSY:
            doc = None
            try:
                # Open the assembly document in Revision Manager
                doc = rm_app.Open(file_path)
                if not doc:
                    logger.warning(f"Revision Manager could not open: {file_path}")
                    return root_node

                linked_docs = doc.LinkedDocuments
                if linked_docs is None:
                    return root_node
                    
                # Deduplicate links by file path and count quantities
                unique_links: Dict[str, Dict[str, Any]] = {}
                
                # Robust extraction of child documents from COM collection
                child_docs_list = []
                try:
                    # Try pythonic iteration first (uses _NewEnum)
                    for c_doc in linked_docs:
                        child_docs_list.append(c_doc)
                except TypeError:
                    # Fallback for late-bound COM objects without _NewEnum
                    count = getattr(linked_docs, "Count", 0)
                    for i in range(1, count + 1):
                        c_doc = None
                        # Try multiple ways to access the item
                        try:
                            c_doc = linked_docs(i)
                        except Exception:
                            try:
                                c_doc = linked_docs.Item(i)
                            except Exception:
                                try:
                                    c_doc = linked_docs[i]
                                except Exception:
                                    pass
                        if c_doc:
                            child_docs_list.append(c_doc)
                
                for child_doc in child_docs_list:
                    try:
                        if not child_doc: continue
                        
                        child_path = child_doc.FullName
                        if not child_path: continue
                            
                        norm_path = os.path.normcase(os.path.abspath(child_path))
                        
                        if norm_path in unique_links:
                            unique_links[norm_path]["quantity"] += 1
                        else:
                            unique_links[norm_path] = {
                                "path": child_path,
                                "quantity": 1
                            }
                    except Exception as inner_e:
                        continue
                
                # Process unique children
                for path, data in unique_links.items():
                    child_path = data["path"]
                    qty = data["quantity"]
                    
                    child_plm_class = self.determine_plm_class(child_path)
                    child_cad = CadFile(file_path=Path(child_path), plm_class=child_plm_class)
                    child_metadata = metadata_service.get_metadata_from_file(child_path)
                    
                    child_node = AssemblyNode(
                        cad_file=child_cad,
                        metadata=child_metadata,
                        level=level + 1,
                        quantity=qty,
                        relationship=RelationshipType.COMPOSED_OF
                    )
                    
                    # Recurse if sub-assembly
                    if child_plm_class == PlmClass.SUB_ASSY:
                        sub_tree = self._explore_with_revision_manager(child_path, rm_app, level + 1)
                        child_node.children = sub_tree.children
                        
                    root_node.add_child(child_node)
                    
            except Exception as e:
                logger.warning(f"Error exploring RM links of {file_path}: {e}")
            # Note: We don't explicitly close the doc as RM handles its own lifecycle 
            # and it might be needed for recursion in some edge cases.

        return root_node

    def _explore_with_solid_edge(self, assembly_doc: Any, level: int = 0) -> AssemblyNode:
        """
        Original exploration using Solid Edge application.
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

        if plm_class == PlmClass.SUB_ASSY or level == 0:
            try:
                if hasattr(assembly_doc, "Occurrences"):
                    occurrences = assembly_doc.Occurrences
                    unique_occs: Dict[str, Dict[str, Any]] = {}
                    
                    for i in range(1, occurrences.Count + 1):
                        occ = occurrences.Item(i)
                        try:
                            occ_path = ""
                            try: occ_path = occ.OccurrenceDocument.FullName
                            except:
                                try: occ_path = occ.OccurrenceFileName
                                except:
                                    try: occ_path = occ.FileName
                                    except: pass
                            
                            if not occ_path: continue
                                
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

                    for path, data in unique_occs.items():
                        occ = data["obj"]
                        occ_path = data["path"]
                        qty = data["quantity"]
                        
                        occ_plm_class = self.determine_plm_class(occ_path)
                        occ_cad = CadFile(file_path=Path(occ_path), plm_class=occ_plm_class)
                        occ_metadata = metadata_service.get_metadata_from_file(occ_path)
                        
                        child_node = AssemblyNode(
                            cad_file=occ_cad,
                            metadata=occ_metadata,
                            level=level + 1,
                            quantity=qty,
                            relationship=RelationshipType.COMPOSED_OF
                        )
                        
                        try:
                            if occ.Subassembly:
                                sub_doc = occ.OccurrenceDocument
                                sub_tree = self._explore_with_solid_edge(sub_doc, level + 1)
                                child_node.children = sub_tree.children
                        except:
                            pass
                            
                        root_node.add_child(child_node)
            except Exception as e:
                logger.warning(f"Error exploring SE occurrences of {file_path}: {e}")

        return root_node

# Global instance
assembly_service = AssemblyService()
