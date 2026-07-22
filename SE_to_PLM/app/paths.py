import sys
import os
import shutil
from pathlib import Path

def get_resource_dir() -> Path:
    """Retourne le chemin d'accès au dossier des ressources (en lecture seule si compilé)."""
    # Si l'application est compilée avec PyInstaller, sys._MEIPASS pointe vers le dossier temporaire d'extraction
    if getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS'):
        return Path(sys._MEIPASS) / "SE_to_PLM" / "ui" / "resources"
    else:
        # En mode développement, on utilise le chemin relatif au fichier actuel
        return Path(__file__).parent.parent / "ui" / "resources"

def get_style_path() -> Path:
    """Retourne le chemin d'accès au fichier de style QSS."""
    if getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS'):
        return Path(sys._MEIPASS) / "SE_to_PLM" / "ui" / "styles" / "style.qss"
    else:
        return Path(__file__).parent.parent / "ui" / "styles" / "style.qss"

def get_writable_config_path(filename: str) -> Path:
    """
    Retourne un chemin d'accès vers un fichier de configuration modifiable.
    - En mode développement : dans le dossier source 'ui/resources'.
    - En mode compilé (EXE) : dans le dossier '%APPDATA%/SE_to_PLM' de l'utilisateur.
      Si le fichier n'existe pas encore dans APPDATA, il est copié depuis les ressources de l'application.
    """
    default_dir = get_resource_dir()
    
    if getattr(sys, 'frozen', False):
        # Chemin dans le dossier AppData/Roaming de l'utilisateur
        appdata_dir = Path(os.environ.get('APPDATA', Path.home())) / "SE_to_PLM"
        appdata_dir.mkdir(parents=True, exist_ok=True)
        config_file = appdata_dir / filename
        
        # Si le fichier n'existe pas encore dans AppData, on y copie le fichier par défaut
        if not config_file.exists():
            default_file = default_dir / filename
            if default_file.exists():
                try:
                    shutil.copy(default_file, config_file)
                except Exception as e:
                    print(f"Erreur lors de la copie du fichier par défaut {filename} vers APPDATA : {e}")
        return config_file
    else:
        return default_dir / filename
