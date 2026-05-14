import sys
import os
from SE_to_PLM.app.constants import DEFAULT_EXPORT_DIR

def bootstrap():
    """Initialise l'environnement de l'application."""
    
    # Création du dossier d'export par défaut
    if not DEFAULT_EXPORT_DIR.exists():
        DEFAULT_EXPORT_DIR.mkdir(parents=True, exist_ok=True)
        
    # Configuration du path pour permettre les imports relatifs si besoin
    # sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

    print("Bootstrapping complete.")

if __name__ == "__main__":
    bootstrap()
