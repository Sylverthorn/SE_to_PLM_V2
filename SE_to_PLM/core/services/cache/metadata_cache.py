from typing import Optional
from SE_to_PLM.core.services.cache.lru_cache import LRUCache
from SE_to_PLM.core.models.metadata import Metadata
from SE_to_PLM.app.constants import METADATA_CACHE_SIZE

class MetadataCache:
    """
    Specialized cache for storing and retrieving Metadata objects indexed by file path.
    """
    def __init__(self, max_size: int = METADATA_CACHE_SIZE):
        self._cache = LRUCache[str, Metadata](max_size=max_size)

    def get_metadata(self, file_path: str) -> Optional[Metadata]:
        return self._cache.get(file_path)

    def set_metadata(self, file_path: str, metadata: Metadata):
        self._cache.set(file_path, metadata)

    def clear(self):
        self._cache.clear()

# Global instance for project-wide metadata caching
metadata_cache = MetadataCache()
