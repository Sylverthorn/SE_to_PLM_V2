import os
from pathlib import Path
from typing import Optional

class PathResolver:
    """
    Utilities for resolving project paths and normalizing network paths.
    """
    
    @staticmethod
    def get_project_root(asm_path: str, levels_up: int = 3) -> str:
        """
        Estimates the project root by going up a few levels from the main assembly.
        """
        try:
            current = Path(asm_path).parent
            for _ in range(levels_up):
                if current.parent == current: # Reached root
                    break
                current = current.parent
            return str(current.absolute())
        except Exception:
            return os.path.dirname(asm_path)

    @staticmethod
    def normalize_network_path(path_str: str) -> str:
        """
        Normalizes paths, specifically avoiding DNS resolution for IP-based UNC paths
        which can be extremely slow.
        """
        if not path_str:
            return ""
            
        # Check for UNC path with IP: \\192.168.x.x\...
        if path_str.startswith("\\\\"):
            parts = path_str[2:].split("\\", 1)
            server = parts[0]
            # If server part looks like an IP (only dots and digits)
            if server.replace('.', '').isdigit():
                # Just fix slashes and return, avoid os.path.normpath which might trigger lookups
                return path_str.replace('/', '\\')
        
        return os.path.normpath(path_str)

# Global instance
path_resolver = PathResolver()
