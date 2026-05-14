import os
from typing import List, Tuple

def scan_directory_unit(dir_path: str, current_depth: int, max_depth: int) -> Tuple[List[Tuple[str, str]], List[str], int]:
    """
    Unit of work for the thread pool. Scans a single directory.
    Returns: (found_dfts, sub_dirs_to_explore, folders_scanned_count)
    """
    found_dfts = []
    sub_dirs = []
    
    try:
        with os.scandir(dir_path) as entries:
            for entry in entries:
                if entry.is_file():
                    if entry.name.lower().endswith('.dft'):
                        # (name_lower_no_ext, full_path)
                        name_no_ext = os.path.splitext(entry.name)[0].lower()
                        found_dfts.append((name_no_ext, entry.path))
                elif entry.is_dir() and current_depth < max_depth:
                    sub_dirs.append(entry.path)
    except (PermissionError, OSError):
        # Silently skip inaccessible directories
        pass
        
    return found_dfts, sub_dirs, 1
