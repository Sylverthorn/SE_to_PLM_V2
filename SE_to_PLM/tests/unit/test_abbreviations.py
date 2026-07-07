import pytest
import os
import json
from pathlib import Path
from SE_to_PLM.core.services.plm.abbreviation_service import AbbreviationService

@pytest.fixture
def temp_abbrev_service(tmp_path):
    """Crée une instance d'AbbreviationService avec un fichier JSON de test temporaire."""
    service = AbbreviationService()
    # Rediriger le fichier json vers un dossier temporaire
    service.json_path = tmp_path / "test_abbreviations.json"
    service.resources_dir = tmp_path
    
    # Remplir avec un échantillon d'abréviations de test
    test_data = [
        {"terme": "Moteur", "abreviation": "MOT", "priorite": 2},
        {"terme": "Afficheur", "abreviation": "AFF", "priorite": 2},
        {"terme": "Boucheuse", "abreviation": "BOUC", "priorite": 1},
        {"terme": "Bouchage", "abreviation": "BOUCH", "priorite": 1},
        {"terme": "Bouchon", "abreviation": "BCH", "priorite": 1},
        {"terme": "Support(s)", "abreviation": "SUPP", "priorite": 2},
        {"terme": "Châssis", "abreviation": "CH", "priorite": 1},
        {"terme": "Écoulement", "abreviation": "ECOU", "priorite": 1},
        {"terme": "Arrière", "abreviation": "AR", "priorite": 2},
        {"terme": "Ensemble", "abreviation": "ENS", "priorite": 1},
    ]
    service.save_abbreviations(test_data)
    return service

def test_majuscules_sans_accents(temp_abbrev_service):
    service = temp_abbrev_service
    assert service.majuscules_sans_accents("Écoulement arrière") == "ECOULEMENT ARRIERE"
    assert service.majuscules_sans_accents("Châssis") == "CHASSIS"
    assert service.majuscules_sans_accents("Moteur") == "MOTEUR"

def test_optimiser_designation_no_change(temp_abbrev_service):
    service = temp_abbrev_service
    # Texte court (< 32 car.) ne contenant aucun terme du dictionnaire
    opt, modified = service.optimiser_designation("Petit Machin")
    assert opt == "Petit Machin"
    assert not modified

def test_optimiser_designation_short_with_abbrev(temp_abbrev_service):
    service = temp_abbrev_service
    # Texte court (< 32 car.) contenant des termes du dictionnaire ne doit PAS être abrégé
    opt, modified = service.optimiser_designation("ENsemble de guidage")
    assert opt == "ENsemble de guidage"
    assert not modified

def test_optimiser_designation_exactly_32_with_abbrev(temp_abbrev_service):
    service = temp_abbrev_service
    # Texte de 32 caractères exacts, ne doit PAS être abrégé
    texte = "Moteur de la boucheuse principal"  # 32 caractères
    opt, modified = service.optimiser_designation(texte)
    assert not modified
    assert opt == "Moteur de la boucheuse principal"

def test_optimiser_designation_strictly_greater_than_32_with_abbrev(temp_abbrev_service):
    service = temp_abbrev_service
    # Texte de 33 caractères, doit être abrégé
    texte = "Moteur de la boucheuse principale"  # 33 caractères
    opt, modified = service.optimiser_designation(texte)
    assert modified
    assert "MOT" in opt
    assert "BOUC" in opt

def test_optimiser_designation_with_abbrev(temp_abbrev_service):
    service = temp_abbrev_service
    # Texte long qui dépasse 32 caractères et peut être abrégé
    texte = "Moteur de la boucheuse principale du châssis"
    # Longueur originale: 44 caractères
    opt, modified = service.optimiser_designation(texte)
    assert modified
    # "Moteur" -> "MOT" (priorité 2)
    # "boucheuse" -> "BOUC" (priorité 1)
    # "châssis" -> "CH" (priorité 1)
    # Et converti en MAJ sans accents car modifié
    assert "MOT" in opt
    assert "BOUC" in opt
    assert "CH" in opt
    assert len(opt) <= 32

def test_optimiser_designation_truncation(temp_abbrev_service):
    service = temp_abbrev_service
    # Texte tellement long que même abrégé, il doit être tronqué
    texte = "Tres Long Texte Sans Rapport Avec Les Mots Du Dictionnaire Mais Quand Meme Tres Long"
    opt, modified = service.optimiser_designation(texte)
    assert modified
    assert len(opt) <= 32
    # Vérifie qu'on n'a pas coupé au milieu d'un mot
    assert opt.endswith("Mots") or opt.endswith("Avec") or opt.endswith("Rapport")

def test_mots_a_preserver(temp_abbrev_service):
    service = temp_abbrev_service
    # "GALAXY" est dans MOTS_A_PRESERVER
    texte = "Support de guidage spécial GALAXY"
    opt, modified = service.optimiser_designation(texte)
    # "GALAXY" doit être préservé à la fin même après tronquage
    assert "GALAXY" in opt
    assert len(opt) <= 32

def test_import_export_xlsx(temp_abbrev_service, tmp_path):
    service = temp_abbrev_service
    export_path = tmp_path / "exported_abbrevs.xlsx"
    
    # Exporter
    service.export_to_excel(str(export_path))
    assert export_path.exists()
    
    # Créer un nouveau service et importer
    new_service = AbbreviationService()
    new_service.json_path = tmp_path / "imported_abbreviations.json"
    new_service.resources_dir = tmp_path
    new_service.save_abbreviations([]) # Initialise vide
    
    nb = new_service.import_from_excel_or_ods(str(export_path), mode="replace")
    assert nb == len(service.abbreviations)
    assert len(new_service.abbreviations) == len(service.abbreviations)
