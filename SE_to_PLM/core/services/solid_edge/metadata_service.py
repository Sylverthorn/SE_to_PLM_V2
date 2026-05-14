from typing import Any, Optional
from SE_to_PLM.core.models.metadata import Metadata
from SE_to_PLM.core.services.cache.metadata_cache import metadata_cache
from SE_to_PLM.core.services.solid_edge.property_reader import property_reader

class MetadataService:
    """
    Coordinates metadata extraction, leveraging caching and specialized property readers.
    """
    
    def get_metadata_from_file(self, file_path: str, use_cache: bool = True) -> Metadata:
        """
        Gets metadata from a file path. Checks cache first.
        """
        if use_cache:
            cached = metadata_cache.get_metadata(file_path)
            if cached:
                return cached
        
        meta = property_reader.extract_from_file(file_path)
        
        if use_cache:
            metadata_cache.set_metadata(file_path, meta)
            
        return meta

    def get_metadata_from_object(self, doc_obj: Any, use_cache: bool = True) -> Metadata:
        """
        Gets metadata from an active COM object. Checks cache using FullName.
        """
        try:
            file_path = doc_obj.FullName
        except:
            return property_reader.extract_from_object(doc_obj)

        if use_cache:
            cached = metadata_cache.get_metadata(file_path)
            if cached:
                return cached
        
        meta = property_reader.extract_from_object(doc_obj)
        
        if use_cache:
            metadata_cache.set_metadata(file_path, meta)
            
        return meta

# Global instance
metadata_service = MetadataService()
