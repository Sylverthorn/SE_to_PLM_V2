import pytest
import time
import threading
from SE_to_PLM.core.services.logging.logger_service import LoggerService, LogLevel
from SE_to_PLM.core.services.cache.lru_cache import LRUCache

def test_logger_callbacks():
    logger = LoggerService()
    received = []
    
    def callback(msg, level):
        received.append((msg, level))
        
    logger.register_callback(callback)
    logger.info("Test Info")
    logger.error("Test Error")
    
    assert len(received) == 2
    assert received[0] == ("Test Info", LogLevel.INFO)
    assert received[1] == ("Test Error", LogLevel.ERROR)

def test_lru_cache_eviction():
    cache = LRUCache[str, int](max_size=2)
    cache.set("a", 1)
    cache.set("b", 2)
    cache.set("c", 3) # Should evict "a"
    
    assert cache.get("a") is None
    assert cache.get("b") == 2
    assert cache.get("c") == 3
    assert len(cache) == 2

def test_lru_cache_access_order():
    cache = LRUCache[str, int](max_size=2)
    cache.set("a", 1)
    cache.set("b", 2)
    cache.get("a") # Access "a", making it MRU
    cache.set("c", 3) # Should evict "b"
    
    assert cache.get("b") is None
    assert cache.get("a") == 1
    assert cache.get("c") == 3

def test_cache_thread_safety():
    cache = LRUCache[int, int](max_size=100)
    
    def worker():
        for i in range(1000):
            cache.set(i % 100, i)
            cache.get(i % 100)
            
    threads = [threading.Thread(target=worker) for _ in range(10)]
    for t in threads: t.start()
    for t in threads: t.join()
    
    assert len(cache) <= 100
