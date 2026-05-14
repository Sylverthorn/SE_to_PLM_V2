import pytest
from SE_to_PLM.core.services.plm.revision_service import RevisionService

def test_revision_service():
    service = RevisionService()
    
    # Standard
    assert service.calculate_previous_indices("C") == ("B", "A")
    assert service.calculate_previous_indices("B") == ("A", "-")
    assert service.calculate_previous_indices("A") == ("-", "-")
    
    # Double letters
    assert service.calculate_previous_indices("AB") == ("AA", "Z")
    assert service.calculate_previous_indices("AA") == ("Z", "-")
    assert service.calculate_previous_indices("BA") == ("AZ", "AY")
    
    # Edge cases
    assert service.calculate_previous_indices("-") == ("-", "-")
    assert service.calculate_previous_indices("") == ("-", "-")
    assert service.calculate_previous_indices("0") == ("-", "-")
