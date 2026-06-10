import sys
import os
from datetime import datetime
from pathlib import Path
from PyQt5.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QLabel, 
    QLineEdit, QPushButton, QComboBox, QProgressBar, QTextEdit, 
    QFileDialog, QMessageBox, QGroupBox, QApplication, QTabWidget,
    QCheckBox, QRadioButton, QButtonGroup, QSplitter, QScrollArea, QSlider, QSpinBox
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
        self.setMinimumSize(900, 800)
        self.resize(1000, 850)
        self._thread = None
        
        self._setup_ui()
        self._load_style()

    def _setup_ui(self):
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        main_layout.setSpacing(10)
        main_layout.setContentsMargins(15, 15, 15, 15)

        # --- Main Splitter (Top content vs Console) ---
        self.splitter = QSplitter(Qt.Vertical)
        
        # --- Top Container (Tabs + Output + Action) ---
        top_container = QWidget()
        top_layout = QVBoxLayout(top_container)
        top_layout.setContentsMargins(0, 0, 0, 0)
        top_layout.setSpacing(10)

        # 1. Tab Widget
        self.tabs = QTabWidget()
        top_layout.addWidget(self.tabs, stretch=1)

        # 1.1. Tab: Assembly
        self.tab_assembly = QWidget()
        self._setup_assembly_tab(self.tab_assembly)
        self.tabs.addTab(self._wrap_in_scroll_area(self.tab_assembly), "Export Assemblage")

        # 1.2. Tab: Batch (PAR/PSM)
        self.tab_batch = QWidget()
        self._setup_batch_tab(self.tab_batch)
        self.tabs.addTab(self._wrap_in_scroll_area(self.tab_batch), "Export par Lot")

        # 1.3. Tab: Multi-ASM (Unified)
        self.tab_multi_asm = QWidget()
        self._setup_multi_asm_tab(self.tab_multi_asm)
        self.tabs.addTab(self._wrap_in_scroll_area(self.tab_multi_asm), "Multi-ASM")

        # 2. Output Group
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
        top_layout.addWidget(group_output)

        # 3. Action & Progress
        action_layout = QVBoxLayout()
        
        buttons_layout = QHBoxLayout()
        self.btn_extraire = QPushButton("LANCER L'EXTRACTION")
        self.btn_extraire.setObjectName("btn_extraire")
        self.btn_extraire.clicked.connect(self._start_extraction)
        buttons_layout.addWidget(self.btn_extraire, stretch=2)

        self.btn_annuler = QPushButton("ANNULER")
        self.btn_annuler.setObjectName("btn_annuler")
        self.btn_annuler.setEnabled(False)
        self.btn_annuler.clicked.connect(self._cancel_extraction)
        buttons_layout.addWidget(self.btn_annuler, stretch=1)
        
        action_layout.addLayout(buttons_layout)

        self.progress_bar = QProgressBar()
        self.progress_bar.setValue(0)
        action_layout.addWidget(self.progress_bar)
        
        self.status_label = QLabel("Prêt")
        action_layout.addWidget(self.status_label)
        top_layout.addLayout(action_layout)

        self.splitter.addWidget(top_container)

        # --- Bottom Container (Console) ---
        bottom_container = QWidget()
        bottom_layout = QVBoxLayout(bottom_container)
        bottom_layout.setContentsMargins(0, 5, 0, 0)
        
        bottom_layout.addWidget(QLabel("Journal d'exécution :"))
        self.console = QTextEdit()
        self.console.setObjectName("console")
        self.console.setReadOnly(True)
        bottom_layout.addWidget(self.console)
        
        self.splitter.addWidget(bottom_container)
        
        # Splitter sizing: give more space to top content by default
        self.splitter.setStretchFactor(0, 3)
        self.splitter.setStretchFactor(1, 1)
        
        main_layout.addWidget(self.splitter)
        
        self.tabs.currentChanged.connect(self._on_tab_changed)

    def _wrap_in_scroll_area(self, widget):
        """Enveloppe un widget dans une zone de défilement pour éviter l'écrasement."""
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QScrollArea.NoFrame)
        scroll.setWidget(widget)
        return scroll

    def _setup_assembly_tab(self, parent_widget):
        layout = QVBoxLayout(parent_widget)
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

        # Options d'indexation (Partagé par convention mais widgets distincts par onglet)
        self.asm_dft_group, self.asm_mode_combo, self.asm_dft_edit, self.asm_btn_dft_browse = self._create_dft_indexing_group()
        layout.addWidget(self.asm_dft_group)
        
        layout.addStretch()

    def _setup_batch_tab(self, parent_widget):
        layout = QVBoxLayout(parent_widget)
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

        # Ajout de la logique DFT pour le lot
        self.batch_dft_group, self.batch_mode_combo, self.batch_dft_edit, self.batch_btn_dft_browse = self._create_dft_indexing_group()
        layout.addWidget(self.batch_dft_group)

        layout.addStretch()

    def _setup_multi_asm_tab(self, parent_widget):
        layout = QVBoxLayout(parent_widget)
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
        self.radio_out_single.setChecked(True)
        out_format_layout.addWidget(self.radio_out_multiple)
        out_format_layout.addWidget(self.radio_out_single)
        layout.addWidget(group_out)

        self.radio_out_multiple.toggled.connect(self._update_multi_asm_labels)
        self.radio_out_single.toggled.connect(self._update_multi_asm_labels)
        
        # Ajout de la logique DFT pour le Multi-ASM
        self.multi_dft_group, self.multi_mode_combo, self.multi_dft_edit, self.multi_btn_dft_browse = self._create_dft_indexing_group()
        layout.addWidget(self.multi_dft_group)

        # 3. Chunk Processing Options
        group_chunk = QGroupBox("Traitement par lots (Chunking)")
        chunk_layout = QVBoxLayout(group_chunk)
        
        # Message d'aide
        self.multi_chunk_help_label = QLabel("Note : Le traitement par lots nécessite le mode de sortie 'Un seul fichier'.")
        self.multi_chunk_help_label.setStyleSheet("color: #888; font-style: italic; margin-bottom: 2px;")
        chunk_layout.addWidget(self.multi_chunk_help_label)
        
        # Ligne avec checkbox et contrôle de taille de lot
        checkbox_size_row = QHBoxLayout()
        self.multi_chunk_enable_checkbox = QCheckBox("Activer le traitement par lots")
        self.multi_chunk_enable_checkbox.setChecked(False)
        self.multi_chunk_enable_checkbox.toggled.connect(self._on_chunk_mode_toggled)
        checkbox_size_row.addWidget(self.multi_chunk_enable_checkbox)
        
        checkbox_size_row.addWidget(QLabel("Taille du lot:"))
        self.multi_chunk_size_spinbox = QSpinBox()
        self.multi_chunk_size_spinbox.setMinimum(10)
        self.multi_chunk_size_spinbox.setMaximum(1000)
        self.multi_chunk_size_spinbox.setValue(100)
        self.multi_chunk_size_spinbox.setSingleStep(10)
        self.multi_chunk_size_spinbox.valueChanged.connect(self._on_chunk_size_changed)
        checkbox_size_row.addWidget(self.multi_chunk_size_spinbox)
        checkbox_size_row.addStretch()
        chunk_layout.addLayout(checkbox_size_row)
        
        # Conteneur pour les contrôles du slider (initially hidden)
        self.multi_chunk_controls_widget = QWidget()
        chunk_controls_layout = QVBoxLayout(self.multi_chunk_controls_widget)
        chunk_controls_layout.setContentsMargins(10, 10, 10, 0)
        
        # Info sur le nombre de fichiers et de lots
        info_row = QHBoxLayout()
        info_row.addWidget(QLabel("Fichiers .asm trouvés:"))
        self.multi_chunk_file_count_label = QLabel("0")
        info_row.addWidget(self.multi_chunk_file_count_label)
        info_row.addStretch()
        chunk_controls_layout.addLayout(info_row)
        
        lot_count_row = QHBoxLayout()
        lot_count_row.addWidget(QLabel("Nombre de lots (taille 100):"))
        self.multi_chunk_count_label = QLabel("1")
        lot_count_row.addWidget(self.multi_chunk_count_label)
        lot_count_row.addStretch()
        chunk_controls_layout.addLayout(lot_count_row)
        
        # Slider pour sélectionner le lot
        slider_row = QHBoxLayout()
        slider_row.addWidget(QLabel("Sélectionner le lot:"))
        self.multi_chunk_slider = QSlider(Qt.Horizontal)
        self.multi_chunk_slider.setMinimum(0)
        self.multi_chunk_slider.setMaximum(0)
        self.multi_chunk_slider.setValue(0)
        self.multi_chunk_slider.setTickPosition(QSlider.TicksBelow)
        self.multi_chunk_slider.setTickInterval(1)
        self.multi_chunk_slider.valueChanged.connect(self._update_multi_chunk_info)
        slider_row.addWidget(self.multi_chunk_slider)
        chunk_controls_layout.addLayout(slider_row)
        
        # Label pour afficher le statut du lot courant
        self.multi_chunk_info_label = QLabel("Lot 1 sur 1 (fichiers 0-0)")
        chunk_controls_layout.addWidget(self.multi_chunk_info_label)
        
        # Initialiser le widget comme invisible
        self.multi_chunk_controls_widget.setVisible(False)
        chunk_layout.addWidget(self.multi_chunk_controls_widget)
        
        layout.addWidget(group_chunk)

        layout.addStretch()

    def _create_dft_indexing_group(self):
        """Crée un groupe d'options pour l'indexation des plans (réutilisable)."""
        group = QGroupBox("Options d'indexation des plans")
        layout = QVBoxLayout(group)
        
        mode_row = QHBoxLayout()
        mode_row.addWidget(QLabel("Mode de recherche des .dft :"))
        mode_combo = QComboBox()
        mode_combo.addItems(["Les deux (Recommandé)", "Arborescence projet uniquement", "Dossier spécifique uniquement"])
        mode_row.addWidget(mode_combo)
        layout.addLayout(mode_row)
        
        dft_row = QHBoxLayout()
        dft_row.addWidget(QLabel("Dossier plans spécifique :"))
        dft_edit = QLineEdit()
        dft_row.addWidget(dft_edit)
        btn_browse = QPushButton("Parcourir...")
        dft_row.addWidget(btn_browse)
        layout.addLayout(dft_row)
        
        # Connexions locales
        mode_combo.currentIndexChanged.connect(lambda idx: self._on_mode_changed_local(idx, dft_edit, btn_browse))
        btn_browse.clicked.connect(lambda: self._browse_dft_folder_local(dft_edit))
        
        # Etat initial
        self._on_mode_changed_local(mode_combo.currentIndex(), dft_edit, btn_browse)
        
        return group, mode_combo, dft_edit, btn_browse

    def _on_mode_changed_local(self, index, dft_edit, btn_browse):
        enabled = (index != 1)
        dft_edit.setEnabled(enabled)
        btn_browse.setEnabled(enabled)

    def _browse_dft_folder_local(self, dft_edit):
        folder = QFileDialog.getExistingDirectory(self, "Dossier plans")
        if folder: dft_edit.setText(folder)

    def _scan_and_update_chunks(self, folder_path):
        """Scanne un dossier pour compter les fichiers .asm et mettre à jour les contrôles de chunk."""
        try:
            asm_count = 0
            for root, _, files in os.walk(folder_path):
                for f in files:
                    if f.lower().endswith(".asm"):
                        asm_count += 1
            
            if asm_count == 0:
                self.multi_chunk_file_count_label.setText("0")
                self.multi_chunk_count_label.setText("0")
                self.multi_chunk_slider.setMaximum(0)
                self.multi_chunk_slider.setValue(0)
                self.multi_chunk_info_label.setText("Aucun fichier .asm trouvé")
                return
            
            chunk_size = self.multi_chunk_size_spinbox.value()
            total_chunks = (asm_count + chunk_size - 1) // chunk_size
            
            self.multi_chunk_file_count_label.setText(str(asm_count))
            self.multi_chunk_count_label.setText(str(total_chunks))
            self.multi_chunk_slider.setMaximum(max(0, total_chunks - 1))
            self.multi_chunk_slider.setValue(0)
            
            self._update_multi_chunk_info()
        except Exception as e:
            self.multi_chunk_info_label.setText(f"Erreur lors du scan: {str(e)}")

    def _update_multi_chunk_info(self):
        """Met à jour l'affichage des infos du chunk courant quand le slider change."""
        chunk_index = self.multi_chunk_slider.value()
        chunk_size = self.multi_chunk_size_spinbox.value()
        total_chunks = self.multi_chunk_slider.maximum() + 1
        
        start_idx = chunk_index * chunk_size
        total_files = int(self.multi_chunk_file_count_label.text()) if self.multi_chunk_file_count_label.text() else 0
        end_idx = min((chunk_index + 1) * chunk_size, total_files)
        
        display_text = f"Lot {chunk_index + 1} sur {total_chunks} (fichiers {start_idx}-{end_idx-1})"
        self.multi_chunk_info_label.setText(display_text)
        
        # Mettre à jour le nom du fichier de sortie avec le numéro du lot
        if self.multi_chunk_enable_checkbox.isChecked():
            current_name = self.output_name_edit.text()
            # Enlever le suffixe de lot existant s'il y en a un
            if " - Lot " in current_name:
                base_name = current_name.split(" - Lot ")[0]
            else:
                base_name = current_name
            # Ajouter le nouveau numéro de lot
            lot_number = f"{chunk_index + 1:03d}"
            self.output_name_edit.setText(f"{base_name} - Lot {lot_number}")

    def _on_chunk_mode_toggled(self, checked):
        """Active/désactive les contrôles de chunking selon l'état de la checkbox."""
        path = self.multi_path_edit.text()
        
        if checked:
            # Activer le mode chunking
            self.multi_chunk_controls_widget.setVisible(True)
            if path and os.path.isdir(path):
                self._scan_and_update_chunks(path)
        else:
            # Désactiver le mode chunking
            self.multi_chunk_controls_widget.setVisible(False)
            # Enlever le suffixe de lot du nom du fichier
            current_name = self.output_name_edit.text()
            if " - Lot " in current_name:
                base_name = current_name.split(" - Lot ")[0]
                self.output_name_edit.setText(base_name)

    def _on_chunk_size_changed(self):
        """Recalcule les chunks quand la taille de lot change."""
        path = self.multi_path_edit.text()
        if path and os.path.isdir(path) and self.multi_chunk_enable_checkbox.isChecked():
            # Recalculer les chunks avec la nouvelle taille
            self._scan_and_update_chunks(path)
            # Réinitialiser le slider à 0
            self.multi_chunk_slider.setValue(0)

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
            if path:
                self.multi_path_edit.setText(path)
                # Désactiver les contrôles de chunk pour un fichier unique
                self.multi_chunk_enable_checkbox.setEnabled(False)
                self.multi_chunk_controls_widget.setVisible(False)
                self._update_multi_asm_labels()
        else:
            path = QFileDialog.getExistingDirectory(self, "Dossier d'assemblages")
            if path:
                self.multi_path_edit.setText(path)
                # Activer la checkbox de chunking pour un dossier
                self.multi_chunk_enable_checkbox.setEnabled(True)
                # Si la checkbox est déjà cochée, scanner le dossier
                if self.multi_chunk_enable_checkbox.isChecked():
                    self._scan_and_update_chunks(path)
                self._update_multi_asm_labels()

    def _update_multi_asm_labels(self):
        path = self.multi_path_edit.text()
        if not path: return
        
        name = Path(path).stem if self.radio_multi_file.isChecked() else Path(path).name
        timestamp = datetime.now().strftime('%Y%m%d_%H%M')
        
        is_folder = self.radio_multi_folder.isChecked()
        is_single_mode = self.radio_out_single.isChecked()
        
        if is_single_mode:
            self.output_label.setText("Nom du fichier/dossier :")
            if is_folder:
                # Mode single avec dossier: les fichiers seront dans un dossier avec les fichiers Lot_001.xlsx, etc. (si chunking)
                self.output_name_edit.setText(f"Export_Multi_ASM_Combined_{name}_{timestamp}")
            else:
                # Mode single avec fichier unique: pas de chunking
                self.output_name_edit.setText(f"Export_Multi_ASM_Combined_{name}_{timestamp}.xlsx")
        else:
            self.output_label.setText("Nom du dossier de sortie :")
            self.output_name_edit.setText(f"Export_Multi_ASM_Files_{name}_{timestamp}")
        
        # Activer/désactiver la checkbox de chunking
        # Le chunking est disponible uniquement quand: dossier + mode single
        should_enable_chunk = is_folder and is_single_mode
        self.multi_chunk_enable_checkbox.setEnabled(should_enable_chunk)
        
        # Si on n'est pas en mode chunking compatible, désactiver le chunking
        if not should_enable_chunk:
            self.multi_chunk_enable_checkbox.setChecked(False)
            self.multi_chunk_controls_widget.setVisible(False)

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
        chunk_index = 0
        chunk_size = -1  # -1 = chunking disabled
        
        # Récupération des widgets selon l'onglet
        if idx == 0:
            input_path = self.input_edit.text()
            mode = "assembly"
            mode_combo = self.asm_mode_combo
            dft_edit = self.asm_dft_edit
        elif idx == 1:
            input_path = self.batch_dir_edit.text()
            mode = "batch"
            recursive = self.cb_recursive.isChecked()
            mode_combo = self.batch_mode_combo
            dft_edit = self.batch_dft_edit
        else:
            input_path = self.multi_path_edit.text()
            mode = "multi_asm"
            is_folder_source = self.radio_multi_folder.isChecked()
            output_mode = "multiple" if self.radio_out_multiple.isChecked() else "single"
            mode_combo = self.multi_mode_combo
            dft_edit = self.multi_dft_edit
            
            # Vérifier si le chunking est activé
            if is_folder_source and output_mode == "single" and self.multi_chunk_enable_checkbox.isChecked():
                # Mode chunking activé: traiter un seul lot
                chunk_index = self.multi_chunk_slider.value()
                chunk_size = self.multi_chunk_size_spinbox.value()

        if not input_path or not os.path.exists(input_path):
            QMessageBox.warning(self, "Erreur", "Vérifiez votre sélection.")
            return

        output_name = self.output_name_edit.text() or "Export_PLM"
        search_mode = [SEARCH_MODE_BOTH, SEARCH_MODE_ARBO, SEARCH_MODE_SPECIFIC][mode_combo.currentIndex()]
        
        self.btn_extraire.setEnabled(False)
        self.btn_extraire.setText("EXTRACTION EN COURS...")
        self.btn_annuler.setEnabled(True)
        self.console.clear()
        
        self._thread = ExtractionThread(
            input_path=input_path, output_dir=str(DEFAULT_EXPORT_DIR), output_name=output_name,
            dft_folder=dft_edit.text(), search_mode=search_mode, mode=mode,
            recursive=recursive, is_folder_source=is_folder_source, output_mode=output_mode,
            chunk_index=chunk_index, chunk_size=chunk_size
        )
        self._thread.progress_signal.connect(self._update_progress)
        self._thread.log_signal.connect(self._log)
        self._thread.finished_signal.connect(self._on_finished)
        self._thread.start()

    def _cancel_extraction(self):
        if self._thread and self._thread.isRunning():
            self.btn_annuler.setEnabled(False)
            self.btn_annuler.setText("ANNULATION...")
            self._thread.cancel()

    def _update_progress(self, val, maximum, msg):
        self.progress_bar.setValue(val)
        self.status_label.setText(msg)

    def _on_finished(self, success, message):
        self.btn_extraire.setEnabled(True)
        self.btn_extraire.setText("LANCER L'EXTRACTION")
        self.btn_annuler.setEnabled(False)
        self.btn_annuler.setText("ANNULER")
        
        if success: QMessageBox.information(self, "Succès", message)
        else: QMessageBox.information(self, "Info", message) if "annulée" in message else QMessageBox.critical(self, "Erreur", message)

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
