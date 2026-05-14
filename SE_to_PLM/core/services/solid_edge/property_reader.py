import win32com.client
import os
from datetime import datetime
from typing import Dict, Any, Optional
from pathlib import Path
from SE_to_PLM.core.models.metadata import Metadata
from SE_to_PLM.core.services.logging.logger_service import logger

class PropertyReader:
    """
    Service for robust metadata extraction from Solid Edge files.
    Optimized for case-insensitivity and multiple property sets.
    """
    
    DATE_FORMATS = [
        "%d/%m/%Y %H:%M:%S", "%d/%m/%Y %H:%M", "%d-%m-%Y %H:%M:%S", "%d-%m-%Y %H:%M",
        "%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M", "%Y/%m/%d %H:%M:%S", "%Y/%m/%d %H:%M",
        "%d/%m/%Y", "%d-%m-%Y", "%d.%m.%Y", "%Y-%m-%d", "%Y/%m/%d",
        "%d/%m/%y", "%d-%m-%y", "%d.%m.%y"
    ]
    TARGET_DATE_FORMAT = "%d/%m/%Y 12:00:00 AM"

    def _get_reader(self):
        try:
            return win32com.client.Dispatch("SolidEdge.FileProperties")
        except Exception as e:
            logger.error(f"Erreur Dispatch FileProperties : {e}")
            return None

    def normalize_date(self, value: Any) -> str:
        """Normalizes a date value (string or pywintypes.datetime) to the PLM target format."""
        if not value:
            return ""
            
        # Handle COM datetime objects
        if hasattr(value, "strftime"):
            return value.strftime(self.TARGET_DATE_FORMAT)
            
        val_str = str(value).strip()
        if not val_str:
            return ""
            
        for fmt in self.DATE_FORMATS:
            try:
                dt = datetime.strptime(val_str, fmt)
                return dt.strftime(self.TARGET_DATE_FORMAT)
            except ValueError:
                continue
        
        # Fallback: simple cleanup
        clean_val = val_str.replace('-', '/').replace('.', '/')
        for fmt in ["%d/%m/%Y", "%Y/%m/%d"]:
            try:
                dt = datetime.strptime(clean_val, fmt)
                return dt.strftime(self.TARGET_DATE_FORMAT)
            except ValueError:
                continue
                
        return val_str

    def _map_property(self, name: str, value: Any, meta: Metadata):
        """Maps a single Solid Edge property to our Metadata model (case-insensitive)."""
        name_lower = name.lower().strip()
        val_str = str(value).strip() if value is not None else ""

        # Designation
        if name_lower in ["désignation", "designation", "desig", "title", "titre"]:
            if not meta.designation: meta.designation = val_str
            
        # Revision / Version
        elif name_lower in ["indice de modification", "revision index", "index", "revision", "rev", "version"]:
            if meta.version == "-": meta.version = val_str
            
        # Author / Creator
        elif name_lower in ["auteur", "author", "username", "créateur", "createur"]:
            if not meta.auteur: meta.auteur = val_str
            
        # Creation Date
        elif name_lower in ["date de création", "date de creation", "creation date", "creationtime"]:
            if not meta.date_creation: meta.date_creation = self.normalize_date(value)
            
        # Modification Author
        elif name_lower in ["auteur modif", "last author", "modifier", "user_version_1"]:
            if not meta.auteur_modif: meta.auteur_modif = val_str
            
        # Modification Date
        elif name_lower in ["date modif", "last modification date", "modification date", "date_version_1"]:
            if not meta.date_modif: meta.date_modif = self.normalize_date(value)
            
        # Material
        elif name_lower in ["matière", "matiere", "material"]:
            if not meta.matiere: meta.matiere = val_str
            
        # Density
        elif name_lower in ["densité", "densite", "density"]:
            if not meta.densite: meta.densite = val_str
            
        # Dia SE
        elif name_lower in ["dia_se", "diamètre", "diametre", "dia"]:
            if not meta.dia_se: meta.dia_se = val_str

    def extract_from_file(self, file_path: str) -> Metadata:
        """Extracts metadata using SolidEdge.FileProperties (No SE UI needed)."""
        meta = Metadata.default()
        reader = self._get_reader()
        if not reader: return meta

        try:
            reader.Open(file_path)
            
            # Scan all property sets: Summary, Extended, Custom
            sets = ["SummaryInformation", "ExtendedSummaryInformation", "Custom"]
            for set_name in sets:
                try:
                    prop_set = reader.Item(set_name)
                    for i in range(1, prop_set.Count + 1):
                        prop = prop_set.Item(i)
                        self._map_property(prop.Name, prop.Value, meta)
                except:
                    pass
            
            # Direct access for sensitive fields (as in working script)
            try:
                custom = reader.Item("Custom")
                for nom_champ, cle_meta in [
                    ("auteur modif",  "auteur_modif"),
                    ("date modif",    "date_modif"),
                    ("Matière",       "matiere"),
                    ("Densité",       "densite"),
                    ("dia_se",        "dia_se"),
                ]:
                    try:
                        p = custom.Item(nom_champ)
                        val = p.Value
                        if val is not None:
                            if "date" in cle_meta:
                                setattr(meta, cle_meta, self.normalize_date(val))
                            else:
                                setattr(meta, cle_meta, str(val).strip())
                    except:
                        pass
            except:
                pass

            reader.Close()
        except Exception:
            logger.warning(f"Lecture propriétés impossible : {os.path.basename(file_path)}")
            try: reader.Close()
            except: pass

        return meta

    def extract_from_object(self, doc_obj: Any) -> Metadata:
        """Extracts metadata from an active COM document object."""
        meta = Metadata.default()
        try:
            prop_sets = doc_obj.Properties
            for i in range(1, prop_sets.Count + 1):
                p_set = prop_sets.Item(i)
                
                # Direct access for sensitive fields if it's the Custom set
                if p_set.Name == "Custom":
                    for nom_champ, cle_meta in [
                        ("auteur modif",  "auteur_modif"),
                        ("date modif",    "date_modif"),
                        ("Matière",       "matiere"),
                        ("Densité",       "densite"),
                        ("dia_se",        "dia_se"),
                    ]:
                        try:
                            p = p_set.Item(nom_champ)
                            val = p.Value
                            if val is not None:
                                if "date" in cle_meta:
                                    setattr(meta, cle_meta, self.normalize_date(val))
                                else:
                                    setattr(meta, cle_meta, str(val).strip())
                        except:
                            pass

                for j in range(1, p_set.Count + 1):
                    prop = p_set.Item(j)
                    self._map_property(prop.Name, prop.Value, meta)
        except Exception as e:
            logger.error(f"Erreur extraction objet COM : {e}")

        return meta

# Global instance
property_reader = PropertyReader()
