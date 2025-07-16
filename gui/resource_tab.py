# gui/resource_tab.py

import os
import traceback
from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QPushButton,
                             QLabel, QProgressBar, QSpacerItem, QSizePolicy,
                             QCheckBox)
from PyQt5.QtCore import Qt, QThread, pyqtSignal

from gui.widgets import FileSelector, StatusPanel
from gui.utils import show_error, get_default_output_path, open_file, show_question

# Import core functionality
import sys

sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from core.excel_handler import ExcelHandler
from core.data_processor import DataProcessor


class ResourceSummaryWorker(QThread):
    """Worker thread for resource summary generation to keep UI responsive"""
    progress_update = pyqtSignal(str)
    finished_signal = pyqtSignal(bool, str, str)

    def __init__(self, input_file, deployments_file, output_file=None):
        super().__init__()
        self.input_file = input_file
        self.deployments_file = deployments_file
        self.output_file = output_file

    def run(self):
        try:
            # Read the input Excel files
            self.progress_update.emit("Reading data from input file...")
            df = ExcelHandler.read_excel(self.input_file)

            self.progress_update.emit("Reading deployments data...")
            deployments_df = ExcelHandler.read_excel(self.deployments_file)

            # Validate the required columns
            self.progress_update.emit("Validating input data...")
            required_columns = ['Ressource', 'Projet', 'Soumise (h)']
            is_valid, missing_columns = DataProcessor.validate_dataframe(df, required_columns)

            if not is_valid:
                error_msg = f"The following required columns are missing: {missing_columns}"
                self.finished_signal.emit(False, error_msg, "")
                return

            # Fix column names if needed (Resource vs Ressource)
            if 'Ressource' not in df.columns and 'Resource' in df.columns:
                df.rename(columns={'Resource': 'Ressource'}, inplace=True)

            # Create lookup dictionaries
            self.progress_update.emit("Creating lookup tables for project information...")
            connection_dict = DataProcessor.create_connection_dict(deployments_df, 'Niveau de connexion')
            phase_dict = DataProcessor.create_connection_dict(deployments_df, 'Phase du projet')
            montant_dict = DataProcessor.create_connection_dict(deployments_df, 'Montant total (Contrat) (Commande)')
            # create CA dictionary by summing CA by project
            ca_dict = {}
            if 'CA' in deployments_df.columns and 'Nom' in deployments_df.columns:
                ca_by_project = deployments_df.groupby('Nom')['CA'].sum()
                ca_dict = ca_by_project.to_dict()

            # Calculate Charge JH
            self.progress_update.emit("Calculating 'Charge JH'...")
            df = DataProcessor.calculate_charge_jh(df)

            # Create pivot table
            self.progress_update.emit("Creating pivot table...")
            pivot_df = ExcelHandler.create_pivot_table(df, 'Charge JH', ['Ressource', 'Projet'])
            # Add CA column to pivot_df by mapping project to summed CA
            pivot_df['CA'] = pivot_df['Projet'].map(ca_dict)

            # Format the resource summary with theoretical charge
            self.progress_update.emit("Formatting output data and calculating theoretical charges...")
            result_df = DataProcessor.format_resource_summary(pivot_df, connection_dict, phase_dict, montant_dict)

            # Get output file path if not specified
            if not self.output_file:
                self.output_file = get_default_output_path(self.input_file, "_resource_summary")

            # Write to Excel
            self.progress_update.emit(f"Writing results to '{self.output_file}'...")
            ExcelHandler.write_excel(result_df, self.output_file, 'Resource Summary')

            self.finished_signal.emit(True, "Resource summary generated successfully!", self.output_file)

        except Exception as e:
            error_msg = f"An error occurred: {str(e)}"
            traceback_str = traceback.format_exc()
            self.finished_signal.emit(False, f"{error_msg}\n\nDetails:\n{traceback_str}", "")


class ResourceSummaryTab(QWidget):
    """Tab for generating resource summary with theoretical charge"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setup_ui()

    def setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(20)

        # Title label
        title_label = QLabel("Générer la consommation de charge")
        title_label.setObjectName("HeaderLabel")
        layout.addWidget(title_label)

        # File selection widgets
        self.input_file_selector = FileSelector("Fichier de consommation:")
        layout.addWidget(self.input_file_selector)

        self.deployments_file_selector = FileSelector("Fichier des déploiements:")
        layout.addWidget(self.deployments_file_selector)

        self.output_file_selector = FileSelector("Fichier de sortie (facultatif):", is_save=True)
        layout.addWidget(self.output_file_selector)

        # Option to open file after generation
        option_layout = QHBoxLayout()
        self.open_after_checkbox = QCheckBox("Ouvrir le fichier après la génération")
        self.open_after_checkbox.setChecked(True)
        option_layout.addWidget(self.open_after_checkbox)
        option_layout.addStretch()
        layout.addLayout(option_layout)

        # Select project phases
        phase_layout = QVBoxLayout()
        phase_label = QLabel("Phases:")
        phase_layout.addWidget(phase_label)

        checkbox_layout = QHBoxLayout()

        checkbox_layout1 = QVBoxLayout()
        checkbox_layout2 = QVBoxLayout()
        checkbox_layout3 = QVBoxLayout()
        checkbox_layout4 = QVBoxLayout()

        self.cadrage_checkbox = QCheckBox("Cadrage / spécification")
        self.cadrage_checkbox.setChecked(True)
        checkbox_layout1.addWidget(self.cadrage_checkbox)

        self.developpement_checkbox = QCheckBox("Développement")
        self.developpement_checkbox.setChecked(True)
        checkbox_layout1.addWidget(self.developpement_checkbox)

        self.production_checkbox = QCheckBox("En production (VSR)")
        self.production_checkbox.setChecked(False)
        checkbox_layout1.addWidget(self.production_checkbox)

        checkbox_layout.addLayout(checkbox_layout1)

        self.non_demarre_autre_lot_checkbox = QCheckBox("Non démarré (autre lot)")
        self.non_demarre_autre_lot_checkbox.setChecked(True)
        checkbox_layout2.addWidget(self.non_demarre_autre_lot_checkbox)

        self.non_demarre_checkbox = QCheckBox("Non démarré (nouveau projet)")
        self.non_demarre_checkbox.setChecked(True)
        checkbox_layout2.addWidget(self.non_demarre_checkbox)

        self.preprod_checkbox = QCheckBox("Pré-production")
        self.preprod_checkbox.setChecked(True)
        checkbox_layout2.addWidget(self.preprod_checkbox)

        checkbox_layout.addLayout(checkbox_layout2)

        self.arrete_checkbox = QCheckBox("Projet arrêté définitivement")
        self.arrete_checkbox.setChecked(False)
        checkbox_layout3.addWidget(self.arrete_checkbox)

        self.pause_checkbox = QCheckBox("Projet en pause")
        self.pause_checkbox.setChecked(True)
        checkbox_layout3.addWidget(self.pause_checkbox)

        self.recette_interne_checkbox = QCheckBox("Recette interne")
        self.recette_interne_checkbox.setChecked(True)
        checkbox_layout3.addWidget(self.recette_interne_checkbox)

        checkbox_layout.addLayout(checkbox_layout3)

        self.recette_user_checkbox = QCheckBox("Recette utilisateur")
        self.recette_user_checkbox.setChecked(True)
        checkbox_layout4.addWidget(self.recette_user_checkbox)

        self.termine_checkbox = QCheckBox("Terminé (VSR signée)")
        self.termine_checkbox.setChecked(False)
        checkbox_layout4.addWidget(self.termine_checkbox)

        checkbox_layout.addLayout(checkbox_layout4)

        phase_layout.addLayout(checkbox_layout)

        phase_layout.addStretch()
        layout.addLayout(phase_layout)

        # Generate button
        button_layout = QHBoxLayout()
        button_layout.addStretch()
        self.generate_button = QPushButton("Générer")
        self.generate_button.clicked.connect(self.generate_summary)
        button_layout.addWidget(self.generate_button)
        layout.addLayout(button_layout)

        # Status panel
        self.status_panel = StatusPanel()
        layout.addWidget(self.status_panel)

        # Progress bar
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        layout.addWidget(self.progress_bar)

        # Add spacer to push everything to the top
        spacer = QSpacerItem(20, 40, QSizePolicy.Minimum, QSizePolicy.Expanding)
        layout.addItem(spacer)

    def generate_summary(self):
        """Start the process to generate resource summary"""
        input_file = self.input_file_selector.get_file_path()
        deployments_file = self.deployments_file_selector.get_file_path()
        output_file = self.output_file_selector.get_file_path()

        # Validate inputs
        if not input_file:
            show_error(self, "Missing Input", "Please select a resource data file.")
            return

        if not deployments_file:
            show_error(self, "Missing Input", "Please select a deployments file.")
            return

        # Disable controls during processing
        self.setEnabled(False)

        # Show progress UI elements
        self.progress_bar.setVisible(True)
        self.progress_bar.setRange(0, 0)  # Indeterminate progress
        self.status_panel.set_status("Processing...", "info")

        # Start worker thread
        self.worker = ResourceSummaryWorker(input_file, deployments_file, output_file)
        self.worker.progress_update.connect(self.update_progress)
        self.worker.finished_signal.connect(self.on_generation_finished)
        self.worker.start()

    def update_progress(self, message):
        """Update progress status with message"""
        self.status_panel.set_status(message, "info")

    def on_generation_finished(self, success, message, output_file):
        """Handle completion of the generation process"""
        # Re-enable controls
        self.setEnabled(True)

        # Update UI
        self.progress_bar.setVisible(False)

        if success:
            self.status_panel.set_status(message, "success")

            # Ask to open the file if checkbox is checked
            if self.open_after_checkbox.isChecked():
                open_file(output_file)
        else:
            self.status_panel.set_status(message, "error")

    def get_checked_phases(self):
        """"
        methode that returns a list containing the phases selected
        """
        phases = []
        if self.cadrage_checkbox.isChecked():
            phases.append("Cadrage / spécification")
        if self.developpement_checkbox.isChecked():
            phases.append("Développement")
        if self.production_checkbox.isChecked():
            phases.append("En production (VSR)")
        if self.non_demarre_autre_lot_checkbox.isChecked():
            phases.append("Non démarré (autre lot)")
        if self.non_demarre_checkbox.isChecked():
            phases.append("Non démarré (nouveau projet)")
        if self.preprod_checkbox.isChecked():
            phases.append("Pré-production")
        if self.arrete_checkbox.isChecked():
            phases.append("Projet arrêté définitivement")
        if self.pause_checkbox.isChecked():
            phases.append("Projet en pause")
        if self.recette_interne_checkbox.isChecked():
            phases.append("Recette interne")
        if self.recette_user_checkbox.isChecked():
            phases.append("Recette utilisateur")
        if self.termine_checkbox.isChecked():
            phases.append("Terminé (VSR signée)")
        return phases
