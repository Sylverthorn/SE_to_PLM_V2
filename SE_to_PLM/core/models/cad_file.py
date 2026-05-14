from dataclasses import dataclass
from pathlib import Path
from SE_to_PLM.core.models.enums import PlmClass

@dataclass
class CadFile:
    file_path: Path
    plm_class: PlmClass
    
    @property
    def name_without_extension(self) -> str:
        return self.file_path.stem
    
    @property
    def extension(self) -> str:
        return self.file_path.suffix.lower()

    @property
    def full_path_str(self) -> str:
        return str(self.file_path.absolute())
