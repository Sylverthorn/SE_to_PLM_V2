from dataclasses import dataclass, field

@dataclass
class Metadata:
    designation: str = ""
    revision: str = "1"
    version: str = "-"
    auteur: str = ""
    date_creation: str = ""
    auteur_modif: str = ""
    date_modif: str = ""
    matiere: str = ""
    densite: str = ""
    dia_se: str = ""
    
    @classmethod
    def default(cls) -> "Metadata":
        return cls()
