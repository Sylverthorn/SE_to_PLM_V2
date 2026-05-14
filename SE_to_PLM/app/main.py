import sys
import os
from pathlib import Path

# Add project root to sys.path to allow absolute imports when running directly
root_path = Path(__file__).parent.parent.parent
if str(root_path) not in sys.path:
    sys.path.insert(0, str(root_path))

from PyQt5.QtWidgets import QApplication
from SE_to_PLM.app.bootstrap import bootstrap
from SE_to_PLM.ui.gui.main_window import MainWindow

def main():
    # 1. Bootstrapping (folders, constants)
    bootstrap()
    
    # 2. Start UI
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()
