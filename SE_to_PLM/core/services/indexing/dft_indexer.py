import os
import time
import concurrent.futures
from typing import Dict, List, Optional, Callable
from SE_to_PLM.core.services.indexing.scan_worker import scan_directory_unit
from SE_to_PLM.core.services.indexing.path_resolver import path_resolver
from SE_to_PLM.core.services.logging.logger_service import logger
from SE_to_PLM.app.constants import SEARCH_MODE_ARBO, SEARCH_MODE_SPECIFIC, SEARCH_MODE_BOTH

class DFTIndexer:
    """
    Parallel drawing indexer service.
    """

    def index_drawings(
        self,
        asm_path: str,
        specific_folder: Optional[str] = None,
        mode: str = SEARCH_MODE_BOTH,
        max_depth: int = 3,
        callback_progress: Optional[Callable[[int, int, int], None]] = None
    ) -> Dict[str, str]:
        """
        Builds an index of {filename_lower: full_path} for all .dft files found.
        """
        start_time = time.time()
        roots_to_scan = []

        # 1. Determine scan roots based on mode
        if mode in [SEARCH_MODE_ARBO, SEARCH_MODE_BOTH]:
            root_arbo = path_resolver.get_project_root(asm_path, levels_up=3)
            roots_to_scan.append(root_arbo)
            logger.info(f"Scan racine projet : {root_arbo}")

        if mode in [SEARCH_MODE_SPECIFIC, SEARCH_MODE_BOTH] and specific_folder:
            roots_to_scan.append(specific_folder)
            logger.info(f"Scan dossier spécifique : {specific_folder}")

        # Remove duplicates and ensure directories exist
        roots_to_scan = list(set([r for r in roots_to_scan if r and os.path.isdir(r)]))
        
        if not roots_to_scan:
            logger.warning("Aucun dossier valide à scanner.")
            return {}

        index: Dict[str, str] = {}
        folders_scanned = 0
        drawings_found = 0
        
        # Initial queue: (path, current_depth)
        queue = [(root, 0) for root in roots_to_scan]
        
        logger.info(f"Lancement du scan parallèle...")

        # 2. Parallel scan with ThreadPoolExecutor
        with concurrent.futures.ThreadPoolExecutor(max_workers=8) as executor:
            while queue:
                # Process queue in batches
                batch = []
                while queue and len(batch) < 16:
                    batch.append(queue.pop(0))
                
                futures = {
                    executor.submit(scan_directory_unit, path, depth, max_depth): (path, depth) 
                    for path, depth in batch
                }

                for future in concurrent.futures.as_completed(futures):
                    try:
                        found_dfts, sub_dirs, scanned_count = future.result()
                        
                        # Add found drawings to index
                        for name, full_path in found_dfts:
                            if name not in index:
                                index[name] = full_path
                                drawings_found += 1
                        
                        folders_scanned += scanned_count
                        
                        # Add sub-directories to queue
                        _, depth = futures[future]
                        for sd in sub_dirs:
                            queue.append((sd, depth + 1))
                            
                        # Progress reporting
                        if callback_progress and folders_scanned % 50 == 0:
                            callback_progress(folders_scanned, drawings_found, 0) # 0 as total (unknown)
                            
                    except Exception as e:
                        logger.error(f"Error scanning directory: {e}")

        duration = time.time() - start_time
        logger.success(f"Indexation terminée : {drawings_found} plans trouvés ({duration:.2f}s).")
        
        return index

# Global instance
dft_indexer = DFTIndexer()
