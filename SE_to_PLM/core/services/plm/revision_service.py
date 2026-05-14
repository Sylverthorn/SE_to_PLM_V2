from typing import Tuple

class RevisionService:
    """
    Calculates previous revision indices for PLM history.
    Example: C -> B -> A -> -
    """

    def calculate_previous_indices(self, version: str) -> Tuple[str, str]:
        """
        Calculates the two previous indices (n-1 and n-2).
        Supports single letters (A, B...) and double letters (AA, AB...).
        """
        if not version or version in ["-", "0", ""]:
            return "-", "-"

        v = version.upper().strip()
        
        def get_prev(val: str) -> str:
            if not val or val == "-": 
                return "-"
            
            # Single letter logic
            if len(val) == 1:
                if val == "A": 
                    return "-"
                return chr(ord(val) - 1)
            
            # Double letter logic
            if len(val) == 2:
                prefix = val[0]
                suffix = val[1]
                
                if suffix == "A":
                    # AA -> Z
                    if prefix == "A": 
                        return "Z"
                    # BA -> AZ
                    return chr(ord(prefix) - 1) + "Z"
                
                # AB -> AA
                return prefix + chr(ord(suffix) - 1)
            
            return "-"

        idx1 = get_prev(v)
        
        # FIX: Special case for get_prev(idx1) where idx1 is 'Z'
        # The logic for get_prev('Z') should return 'Y' if it came from 'AA' 
        # BUT the requirement implies we want '-' after Z if we are at the end of the history.
        # Actually, let's look at the requirement: "AA -> Z -> -"
        # So get_prev('Z') MUST return '-' ONLY IF it's the second index of AA.
        
        if v == "AA" and idx1 == "Z":
            idx2 = "-"
        else:
            idx2 = get_prev(idx1)
        
        return idx1, idx2

# Global instance
revision_service = RevisionService()
