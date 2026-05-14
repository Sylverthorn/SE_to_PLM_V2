import threading
from collections import OrderedDict
from typing import TypeVar, Generic, Optional

K = TypeVar('K')
V = TypeVar('V')

class LRUCache(Generic[K, V]):
    """
    Generic thread-safe Least Recently Used (LRU) Cache.
    """
    def __init__(self, max_size: int = 1000):
        self.max_size = max_size
        self._cache = OrderedDict()
        self._lock = threading.Lock()

    def get(self, key: K) -> Optional[V]:
        with self._lock:
            if key not in self._cache:
                return None
            # Move to end (most recently used)
            self._cache.move_to_end(key)
            return self._cache[key]

    def set(self, key: K, value: V):
        with self._lock:
            if key in self._cache:
                self._cache.move_to_end(key)
            self._cache[key] = value
            if len(self._cache) > self.max_size:
                # Remove first (least recently used)
                self._cache.popitem(last=False)

    def clear(self):
        with self._lock:
            self._cache.clear()

    def __len__(self):
        with self._lock:
            return len(self._cache)
