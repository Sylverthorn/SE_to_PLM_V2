from dataclasses import dataclass
from typing import Any, Dict

@dataclass
class ExportRow:
    level: int               # 1: Level
    relationship: str        # 2: Relationship
    order: int               # 3: ordre
    quantity: int            # 4: quantite
    repere: str              # 5: repere
    special_cad: str         # 6: SpecialCAD
    plm_class: str           # 7: Class
    ref_utilisat: str        # 8: ref_utilisat
    version: str             # 9: version
    indice_1: str            # 10: indice_1
    indice_2: str            # 11: indice_2
    revision: str            # 12: revision
    designation: str         # 13: designation
    cus_createur: str        # 14: cus_createur
    cus_date_crea: str       # 15: cus_date_crea
    user_version_1: str      # 16: user_version_1
    date_version_1: str      # 17: date_version_1
    matiere: str             # 18: matiere
    densite: str             # 19: densite
    dia_se: str              # 20: dia_se
    mode_appro: str          # 21: mode_appro
    attachments: str         # 22: Attachments
    custom_properties: dict = None

    def __post_init__(self):
        # We store the standard fields in special_values so they can be looked up
        self.special_values = {
            "level": self.level,
            "relationship": self.relationship,
            "order": self.order,
            "quantity": self.quantity,
            "repere": self.repere,
            "special_cad": self.special_cad,
            "plm_class": self.plm_class,
            "ref_utilisat": self.ref_utilisat,
            "version": self.version,
            "indice_1": self.indice_1,
            "indice_2": self.indice_2,
            "revision": self.revision,
            "designation": self.designation,
            "cus_createur": self.cus_createur,
            "cus_date_crea": self.cus_date_crea,
            "user_version_1": self.user_version_1,
            "date_version_1": self.date_version_1,
            "matiere": self.matiere,
            "densite": self.densite,
            "dia_se": self.dia_se,
            "mode_appro": self.mode_appro,
            "attachments": self.attachments
        }
        
        # We store all properties in properties dict
        self.properties = {}
        if self.custom_properties:
            for k, v in self.custom_properties.items():
                self.properties[k.lower().strip()] = v
                
        # Also sync standard properties into the properties dictionary for fallback
        for k in ["designation", "revision", "version", "matiere", "densite", "dia_se", 
                  "cus_createur", "cus_date_crea", "user_version_1", "date_version_1"]:
            val = self.special_values[k]
            if val:
                # Add typical mappings
                if k == "matiere":
                    self.properties["matière"] = val
                    self.properties["matiere"] = val
                    self.properties["material"] = val
                elif k == "densite":
                    self.properties["densité"] = val
                    self.properties["densite"] = val
                    self.properties["density"] = val
                elif k == "cus_createur":
                    self.properties["auteur"] = val
                    self.properties["author"] = val
                    self.properties["cus_createur"] = val
                elif k == "cus_date_crea":
                    self.properties["date de création"] = val
                    self.properties["date_creation"] = val
                    self.properties["cus_date_crea"] = val
                elif k == "user_version_1":
                    self.properties["auteur modif"] = val
                    self.properties["user_version_1"] = val
                elif k == "date_version_1":
                    self.properties["date modif"] = val
                    self.properties["date_version_1"] = val
                else:
                    self.properties[k] = val

    def get_value(self, col_def: dict) -> Any:
        source_type = col_def["source_type"]
        source_name = col_def["source_name"]
        default_val = col_def.get("default_value", "")
        
        if source_type == "special_processed":
            attr_map = {
                "level": "level",
                "relationship": "relationship",
                "order": "order",
                "quantity": "quantity",
                "repere": "repere",
                "special_cad": "special_cad",
                "plm_class": "plm_class",
                "ref_utilisat": "ref_utilisat",
                "version": "version",
                "indice_1": "indice_1",
                "indice_2": "indice_2",
                "revision": "revision",
                "designation": "designation",
                "mode_appro": "mode_appro",
                "attachments": "attachments"
            }
            attr_name = attr_map.get(source_name)
            if attr_name and hasattr(self, attr_name):
                return getattr(self, attr_name)
            return self.special_values.get(source_name, default_val)
            
        elif source_type == "solid_edge_property":
            aliases = source_name if isinstance(source_name, list) else [source_name]
            for alias in aliases:
                key = alias.lower().strip()
                # Check dynamic mappings first
                attr_map = {
                    "designation": "designation",
                    "auteur": "cus_createur",
                    "author": "cus_createur",
                    "createur": "cus_createur",
                    "date_creation": "cus_date_crea",
                    "date de création": "cus_date_crea",
                    "auteur_modif": "user_version_1",
                    "auteur modif": "user_version_1",
                    "date_modif": "date_version_1",
                    "date modif": "date_version_1",
                    "matiere": "matiere",
                    "matière": "matiere",
                    "densite": "densite",
                    "densité": "densite",
                    "dia_se": "dia_se",
                    "diametre": "dia_se",
                    "diamètre": "dia_se",
                }
                attr_name = attr_map.get(key)
                if attr_name and hasattr(self, attr_name):
                    return getattr(self, attr_name)
                if hasattr(self, key):
                    return getattr(self, key)
                if key in self.properties:
                    return self.properties[key]
            return default_val
        return default_val

    def to_list(self) -> list:
        """Returns the row as a list for openpyxl (maintained for backward compatibility)."""
        return [
            self.level, self.relationship, self.order, self.quantity,
            self.repere, self.special_cad, self.plm_class, self.ref_utilisat,
            self.version, self.indice_1, self.indice_2, self.revision,
            self.designation, self.cus_createur, self.cus_date_crea,
            self.user_version_1, self.date_version_1, self.matiere,
            self.densite, self.dia_se, self.mode_appro, self.attachments
        ]
