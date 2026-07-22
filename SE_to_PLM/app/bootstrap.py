import sys
import os
import json
from pathlib import Path
from SE_to_PLM.app.constants import DEFAULT_EXPORT_DIR

def bootstrap():
    """Initialise l'environnement de l'application."""
    
    # Création du dossier d'export par défaut
    if not DEFAULT_EXPORT_DIR.exists():
        DEFAULT_EXPORT_DIR.mkdir(parents=True, exist_ok=True)
        
    from SE_to_PLM.app.paths import get_writable_config_path
    
    # Initialisation du dictionnaire d'abréviations
    json_path = get_writable_config_path("abbreviations.json")
    if not json_path.exists():
        json_path.parent.mkdir(parents=True, exist_ok=True)
        try:
            # Recherche du script à la racine
            root_dir = Path(__file__).parent.parent.parent
            sys.path.append(str(root_dir))
            import designation_plm_32_priorite_abreviations_COMPLET_V3_ABREVIATIONS_MAJ as default_script
            
            abbrev_list = []
            seen = set()
            for term, abbrev, priority in default_script.abreviations_source:
                key = term.strip().lower()
                if key not in seen:
                    seen.add(key)
                    abbrev_list.append({
                        "terme": term,
                        "abreviation": abbrev,
                        "priorite": priority
                    })
            if "cache" not in seen:
                abbrev_list.append({
                    "terme": "cache",
                    "abreviation": "cach",
                    "priorite": 2
                })
            with open(json_path, "w", encoding="utf-8") as f:
                json.dump(abbrev_list, f, ensure_ascii=False, indent=2)
            print(f"Dictionnaire d'abréviations initialisé avec {len(abbrev_list)} entrées.")
        except Exception as e:
            with open(json_path, "w", encoding="utf-8") as f:
                json.dump([], f)
            print(f"Dictionnaire d'abréviations initialisé vide (erreur : {e})")

    # Charger les abréviations après initialisation
    try:
        from SE_to_PLM.core.services.plm.abbreviation_service import abbreviation_service
        abbreviation_service.load_abbreviations()
    except Exception as e:
        print(f"Erreur de chargement du service d'abréviation : {e}")

    print("Bootstrapping complete.")

if __name__ == "__main__":
    bootstrap()
