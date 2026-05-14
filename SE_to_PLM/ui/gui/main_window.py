import sys
import os
from datetime import datetime
from pathlib import Path
from PyQt5.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QLabel, 
    QLineEdit, QPushButton, QComboBox, QProgressBar, QTextEdit, 
    QFileDialog, QMessageBox, QGroupBox, QApplication, QTabWidget,
    QCheckBox, QRadioButton, QButtonGroup
)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QTextCursor

from SE_to_PLM.ui.gui.extraction_thread import ExtractionThread
from SE_to_PLM.app.constants import DEFAULT_EXPORT_DIR, SEARCH_MODE_BOTH, SEARCH_MODE_ARBO, SEARCH_MODE_SPECIFIC
from SE_to_PLM.infrastructure.solid_edge.connection_manager import connection_manager

class MainWindow(QMainWindow):
    """
    Fenêtre principale de l'application SE_to_PLM.
    Onglets unifiés pour l'Assemblage, le Lot et le Multi-ASM.
    """
    def __init__(self):
        super().__init__()
        self.setWindowTitle("SE_to_PLM — Export Industriel PLM")
        self.setMinimumSize(850, 750)
        self._thread = None
        
        self._setup_ui()
        self._load_style()

    def _setup_ui(self):
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        main_layout.setSpacing(10)
        main_layout.setContentsMargins(20, 20, 20, 20)

        # --- Tab Widget ---
        self.tabs = QTabWidget()
        main_layout.addWidget(self.tabs)

        # 1. Tab: Assembly
        self.tab_assembly = QWidget()
        self._setup_assembly_tab()
        self.tabs.addTab(self.tab_assembly, "Export Assemblage")

        # 2. Tab: Batch (PAR/PSM)
        self.tab_batch = QWidget()
        self._setup_batch_tab()
        self.tabs.addTab(self.tab_batch, "Export par Lot")

        # 3. Tab: Multi-ASM (Unified)
        self.tab_multi_asm = QWidget()
        self._setup_multi_asm_tab()
        self.tabs.addTab(self.tab_multi_asm, "Multi-ASM")

        # --- Common Sections (Bottom) ---
        
        # Output Group
        group_output = QGroupBox("Destination de l'export")
        out_layout = QVBoxLayout(group_output)
        
        name_row = QHBoxLayout()
        self.output_label = QLabel("Nom du fichier Excel :")
        name_row.addWidget(self.output_label)
        self.output_name_edit = QLineEdit()
        self.output_name_edit.setPlaceholderText("Export_PLM_PROJET.xlsx")
        name_row.addWidget(self.output_name_edit)
        out_layout.addLayout(name_row)
        
        self.dest_label = QLabel(f"Dossier : {DEFAULT_EXPORT_DIR}")
        out_layout.addWidget(self.dest_label)
        main_layout.addWidget(group_output)

        # Action Button
        self.btn_extraire = QPushButton("LANCER L'EXTRACTION")
        self.btn_extraire.setObjectName("btn_extraire")
        self.btn_extraire.clicked.connect(self._start_extraction)
        main_layout.addWidget(self.btn_extraire)

        # Progress
        self.progress_bar = QProgressBar()
        self.progress_bar.setValue(0)
        main_layout.addWidget(self.progress_bar)
        
        self.status_label = QLabel("Prêt")
        main_layout.addWidget(self.status_label)

        # Console
        main_layout.addWidget(QLabel("Journal d'exécution :"))
        self.console = QTextEdit()
        self.console.setObjectName("console")
        self.console.setReadOnly(True)
        main_layout.addWidget(self.console)
        
        self.tabs.currentChanged.connect(self._on_tab_changed)

    def _setup_assembly_tab(self):
        layout = QVBoxLayout(self.tab_assembly)
        layout.setSpacing(15)

        group_selection = QGroupBox("Sélection du projet")
        sel_layout = QVBoxLayout(group_selection)
        file_row = QHBoxLayout()
        file_row.addWidget(QLabel("Fichier principal (.asm, .par) :"))
        self.input_edit = QLineEdit()
        self.input_edit.setReadOnly(True)
        file_row.addWidget(self.input_edit)
        btn_browse = QPushButton("Parcourir...")
        btn_browse.clicked.connect(self._browse_file)
        file_row.addWidget(btn_browse)
        sel_layout.addLayout(file_row)
        layout.addWidget(group_selection)

        group_options = QGroupBox("Options d'indexation des plans")
        opt_layout = QVBoxLayout(group_options)
        mode_row = QHBoxLayout()
        mode_row.addWidget(QLabel("Mode de recherche des .dft :"))
        self.mode_combo = QComboBox()
        self.mode_combo.addItems(["Les deux (Recommandé)", "Arborescence projet uniquement", "Dossier spécifique uniquement"])
        self.mode_combo.currentIndexChanged.connect(self._on_mode_changed)
        mode_row.addWidget(self.mode_combo)
        opt_layout.addLayout(mode_row)
        dft_row = QHBoxLayout()
        dft_row.addWidget(QLabel("Dossier plans spécifique :"))
        self.dft_edit = QLineEdit()
        dft_row.addWidget(self.dft_edit)
        self.btn_dft_browse = QPushButton("Parcourir...")
        self.btn_dft_browse.clicked.connect(self._browse_dft_folder)
        dft_row.addWidget(self.btn_dft_browse)
        opt_layout.addLayout(dft_row)
        layout.addWidget(group_options)
        layout.addStretch()

    def _setup_batch_tab(self):
        layout = QVBoxLayout(self.tab_batch)
        layout.setSpacing(15)
        group_batch = QGroupBox("Sélection du dossier")
        batch_layout = QVBoxLayout(group_batch)
        dir_row = QHBoxLayout()
        dir_row.addWidget(QLabel("Dossier à scanner :"))
        self.batch_dir_edit = QLineEdit()
        self.batch_dir_edit.setReadOnly(True)
        dir_row.addWidget(self.batch_dir_edit)
        btn_browse_dir = QPushButton("Parcourir...")
        btn_browse_dir.clicked.connect(self._browse_batch_directory)
        dir_row.addWidget(btn_browse_dir)
        batch_layout.addLayout(dir_row)
        self.cb_recursive = QCheckBox("Inclure les sous-dossiers")
        batch_layout.addWidget(self.cb_recursive)
        layout.addWidget(group_batch)
        layout.addStretch()

    def _setup_multi_asm_tab(self):
        layout = QVBoxLayout(self.tab_multi_asm)
        layout.setSpacing(15)

        # 1. Source Type
        group_src = QGroupBox("Source")
        src_layout = QVBoxLayout(group_src)
        
        radio_row = QHBoxLayout()
        self.radio_multi_file = QRadioButton("Fichier .asm unique")
        self.radio_multi_folder = QRadioButton("Dossier complet (scan des .asm)")
        self.radio_multi_file.setChecked(True)
        radio_row.addWidget(self.radio_multi_file)
        radio_row.addWidget(self.radio_multi_folder)
        src_layout.addLayout(radio_row)
        
        path_row = QHBoxLayout()
        self.multi_path_edit = QLineEdit()
        self.multi_path_edit.setReadOnly(True)
        path_row.addWidget(self.multi_path_edit)
        btn_browse_multi = QPushButton("Parcourir...")
        btn_browse_multi.clicked.connect(self._browse_multi_path)
        path_row.addWidget(btn_browse_multi)
        src_layout.addLayout(path_row)
        layout.addWidget(group_src)

        # 2. Output Format
        group_out = QGroupBox("Format de sortie")
        out_format_layout = QVBoxLayout(group_out)
        self.radio_out_multiple = QRadioButton("Plusieurs fichiers (Un par sous-assemblage)")
        self.radio_out_single = QRadioButton("Un seul fichier (Tous les blocs Niv 0/1)")
        self.radio_out_multiple.setChecked(True)
        out_format_layout.addWidget(self.radio_out_multiple)
        out_format_layout.addWidget(self.radio_out_single)
        layout.addWidget(group_out)

        self.radio_out_multiple.toggled.connect(self._update_multi_asm_labels)
        self.radio_out_single.toggled.connect(self._update_multi_asm_labels)
        
        layout.addStretch()

    def _load_style(self):
        style_path = Path(__file__).parent.parent / "styles" / "style.qss"
        if style_path.exists():
            with open(style_path, "r", encoding="utf-8") as f:
                self.setStyleSheet(f.read())

    def _browse_file(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "CAO", "", "Solid Edge (*.asm *.par *.psm)")
        if file_path:
            self.input_edit.setText(file_path)
            self.output_name_edit.setText(f"Export_PLM_{Path(file_path).stem}_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx")

    def _browse_batch_directory(self):
        directory = QFileDialog.getExistingDirectory(self, "Dossier")
        if directory:
            self.batch_dir_edit.setText(directory)
            self.output_name_edit.setText(f"Export_Lot_{Path(directory).name}_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx")

    def _browse_multi_path(self):
        if self.radio_multi_file.isChecked():
            path, _ = QFileDialog.getOpenFileName(self, "Assemblage", "", "Assemblage Solid Edge (*.asm)")
        else:
            path = QFileDialog.getExistingDirectory(self, "Dossier d'assemblages")
        
        if path:
            self.multi_path_edit.setText(path)
            self._update_multi_asm_labels()

    def _update_multi_asm_labels(self):
        path = self.multi_path_edit.text()
        if not path: return
        
        name = Path(path).stem if self.radio_multi_file.isChecked() else Path(path).name
        timestamp = datetime.now().strftime('%Y%m%d_%H%M')
        
        if self.radio_out_single.isChecked():
            self.output_label.setText("Nom du fichier Excel :")
            self.output_name_edit.setText(f"Export_Multi_ASM_Combined_{name}_{timestamp}.xlsx")
        else:
            self.output_label.setText("Nom du dossier de sortie :")
            self.output_name_edit.setText(f"Export_Multi_ASM_Files_{name}_{timestamp}")

    def _browse_dft_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "Dossier plans")
        if folder: self.dft_edit.setText(folder)

    def _on_mode_changed(self, index):
        enabled = (index != 1)
        self.dft_edit.setEnabled(enabled)
        self.btn_dft_browse.setEnabled(enabled)

    def _on_tab_changed(self, index):
        if index == 2: # Multi-ASM
            self._update_multi_asm_labels()
        else:
            self.output_label.setText("Nom du fichier Excel :")
            self.output_name_edit.setPlaceholderText("Export_PLM_PROJET.xlsx")

    def _log(self, message: str, level_name: str):
        color = {"success": "#6A9955", "error": "#F44336", "warning": "#CE9178"}.get(level_name, "#d4d4d4")
        formatted = f'<span style="color:{color}">[{datetime.now().strftime("%H:%M:%S")}] {message}</span>'
        self.console.append(formatted)
        self.console.moveCursor(QTextCursor.End)

    def _start_extraction(self):
        idx = self.tabs.currentIndex()
        recursive = False
        is_folder_source = False
        output_mode = "single"
        
        if idx == 0:
            input_path = self.input_edit.text()
            mode = "assembly"
        elif idx == 1:
            input_path = self.batch_dir_edit.text()
            mode = "batch"
            recursive = self.cb_recursive.isChecked()
        else:
            input_path = self.multi_path_edit.text()
            mode = "multi_asm"
            is_folder_source = self.radio_multi_folder.isChecked()
            output_mode = "multiple" if self.radio_out_multiple.isChecked() else "single"

        if not input_path or not os.path.exists(input_path):
            QMessageBox.warning(self, "Erreur", "Vérifiez votre sélection.")
            return

        output_name = self.output_name_edit.text() or "Export_PLM"
        search_mode = [SEARCH_MODE_BOTH, SEARCH_MODE_ARBO, SEARCH_MODE_SPECIFIC][self.mode_combo.currentIndex()]
        
        self.btn_extraire.setEnabled(False)
        self.btn_extraire.setText("EXTRACTION EN COURS...")
        self.console.clear()
        
        self._thread = ExtractionThread(
            input_path=input_path, output_dir=str(DEFAULT_EXPORT_DIR), output_name=output_name,
            dft_folder=self.dft_edit.text(), search_mode=search_mode, mode=mode,
            recursive=recursive, is_folder_source=is_folder_source, output_mode=output_mode
        )
        self._thread.progress_signal.connect(self._update_progress)
        self._thread.log_signal.connect(self._log)
        self._thread.finished_signal.connect(self._on_finished)
        self._thread.start()

    def _update_progress(self, val, maximum, msg):
        self.progress_bar.setValue(val)
        self.status_label.setText(msg)

    def _on_finished(self, success, message):
        self.btn_extraire.setEnabled(True)
        self.btn_extraire.setText("LANCER L'EXTRACTION")
        if success: QMessageBox.information(self, "Succès", message)
        else: QMessageBox.critical(self, "Erreur", message)

    def closeEvent(self, event):
        reply = QMessageBox.question(
            self, 
            "Quitter l'application", 
            "Souhaitez-vous également fermer Solid Edge ?", 
            QMessageBox.Yes | QMessageBox.No | QMessageBox.Cancel, 
            QMessageBox.No
        )
        
        if reply == QMessageBox.Cancel:
            event.ignore()
        elif reply == QMessageBox.Yes:
            # L'utilisateur a choisi de fermer Solid Edge
            connection_manager.quit()
            event.accept()
        else:
            # L'utilisateur ferme juste l'appli
            event.accept()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())
