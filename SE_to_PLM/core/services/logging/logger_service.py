import threading
from typing import Callable, List, Optional
from enum import Enum

class LogLevel(Enum):
    INFO = "info"
    SUCCESS = "success"
    WARNING = "warning"
    ERROR = "error"

class LoggerService:
    """
    Centralized logging service with support for multiple sinks, 
    including UI callbacks and standard output.
    """
    def __init__(self):
        self._callbacks: List[Callable[[str, LogLevel], None]] = []
        self._lock = threading.Lock()

    def register_callback(self, callback: Callable[[str, LogLevel], None]):
        with self._lock:
            self._callbacks.append(callback)

    def unregister_callback(self, callback: Callable[[str, LogLevel], None]):
        with self._lock:
            if callback in self._callbacks:
                self._callbacks.remove(callback)

    def log(self, message: str, level: LogLevel = LogLevel.INFO):
        with self._lock:
            # Print to stdout for console visibility
            print(f"[{level.value.upper()}] {message}")
            
            # Notify all registered callbacks (e.g., UI console)
            for callback in self._callbacks:
                try:
                    callback(message, level)
                except Exception as e:
                    print(f"Error in logger callback: {e}")

    def info(self, message: str):
        self.log(message, LogLevel.INFO)

    def success(self, message: str):
        self.log(message, LogLevel.SUCCESS)

    def warning(self, message: str):
        self.log(message, LogLevel.WARNING)

    def error(self, message: str):
        self.log(message, LogLevel.ERROR)

# Global instance for easy access across the application
logger = LoggerService()
