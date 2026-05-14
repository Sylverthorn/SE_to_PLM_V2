from dataclasses import dataclass

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
    attachments: str         # 21: Attachments

    def to_list(self) -> list:
        """Returns the row as a list for openpyxl."""
        return [
            self.level, self.relationship, self.order, self.quantity,
            self.repere, self.special_cad, self.plm_class, self.ref_utilisat,
            self.version, self.indice_1, self.indice_2, self.revision,
            self.designation, self.cus_createur, self.cus_date_crea,
            self.user_version_1, self.date_version_1, self.matiere,
            self.densite, self.dia_se, self.attachments
        ]
