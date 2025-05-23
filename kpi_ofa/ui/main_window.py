# kpi_ofa/ui/main_window.py

import os
import time
import pandas as pd
from typing import Dict, Tuple

from PyQt5.QtWidgets import (QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
                            QLabel, QLineEdit, QPushButton, QStatusBar, 
                            QMessageBox, QGroupBox, QSizePolicy, QProgressBar,
                            QFileDialog)
from PyQt5.QtCore import Qt, QDate

from datetime import datetime

from kpi_ofa.ui.widgets.log_widget import LogWidget
from kpi_ofa.ui.widgets.date_widget import DateRangeWidget
from kpi_ofa.ui.config_dialog import ConfigDialog
from kpi_ofa.core.excel_data_processor import ExcelDataProcessor
from kpi_ofa.core.log_manager import LogManager
from kpi_ofa.services.sap_connection import SAPGuiConnection
from kpi_ofa.services.sap_transactions import SAPDataExtractor

import kpi_ofa.constants as constants

import logging
logger = logging.getLogger(__name__)

class MainWindow(QMainWindow):
    """
    Finestra principale dell'applicazione KPI OFA.
    
    Questa classe è responsabile della creazione e gestione dell'interfaccia
    utente principale e della coordinazione tra i vari componenti dell'applicazione.
    """
    
    def __init__(self, config_manager):
        """
        Inizializza la finestra principale dell'applicazione.
        
        Args:
            config_manager: Gestore della configurazione dell'applicazione.
        """
        super().__init__()
        
        # Salva il gestore della configurazione
        self.config_manager = config_manager
        
        # Ottieni l'istanza del LogManager
        self.log_manager = LogManager()

        # Configura la finestra
        self.setup_window()    

        # Crea l'interfaccia utente
        self.setup_ui()
        
        # Inizializza i componenti di supporto
        self.init_components()
        
        # Configura i segnali
        self.setup_signals()
        
        # Carica la configurazione
        self.load_config()
        
        # Messaggio di avvio
        self.log_manager.log("Sistema pronto!", "info", origin=logger.name)
    
    def init_components(self):
        """Inizializza i componenti di supporto dell'applicazione."""
        # Configura i callback per il LogManager
        self.log_manager.set_ui_callback(self.on_log_message)
        self.log_manager.set_status_bar_callback(self.update_status_bar)
        
        # Processore di dati
        self.excel_data_processor = ExcelDataProcessor()
        
        # Factory per le connessioni SAP
        self.sap_factory = SAPGuiConnection()
        
        # Percorso del file Excel selezionato
        self.excel_file_path = None
    
    def setup_window(self):
        """Configura le proprietà della finestra."""
        self.setWindowTitle("Estrai Dati KPI OFA")
        self.setWindowIconText("KPI OFA")
        self.setGeometry(100, 100, 800, 600)  # x, y, width, height
    
    def setup_ui(self):
        """Crea l'interfaccia utente della finestra principale."""
        # Widget centrale
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        # Layout principale
        main_layout = QVBoxLayout(central_widget)
        
        # Layout orizzontale per i due pannelli
        content_layout = QHBoxLayout()
        
        # Pannello sinistro (controlli)
        left_panel = self.create_left_panel()
        content_layout.addWidget(left_panel)
        # Imposta l'allineamento in alto per il gruppo di sinistra
        content_layout.setAlignment(left_panel, Qt.AlignTop)
        
        # Pannello destro (log)
        right_panel = self.create_right_panel()
        content_layout.addWidget(right_panel)
        
        # Barra dei pulsanti inferiore
        button_bar = self.create_button_bar()
        
        # Aggiungi i layout al layout principale
        main_layout.addLayout(content_layout)
        main_layout.addWidget(button_bar)
        
        # Status Bar
        self.setup_status_bar()
    
    def create_left_panel(self):
        """
        Crea il pannello sinistro con i controlli.
        
        Returns:
            QGroupBox: Il pannello sinistro.
        """
        left_group = QGroupBox("Estrazioni SAP")
        left_layout = QVBoxLayout(left_group)
        
        # Imposta le dimensioni
        left_group.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
        left_group.setFixedWidth(250)
        
        # Widget per la selezione delle date
        self.date_widget = DateRangeWidget()
        left_layout.addWidget(self.date_widget)
        
        # Sezione per la selezione del file Excel
        file_layout = QVBoxLayout()
        
        # Campo di testo di sola lettura per mostrare il nome del file
        self.file_text = QLineEdit()
        self.file_text.setReadOnly(True)
        self.file_text.setPlaceholderText("Nessun file selezionato...")
        
        # Pulsante per selezionare il file
        self.browse_button = QPushButton("Seleziona Excel...")
        self.browse_button.clicked.connect(self.select_excel_file)
        
        # Aggiungi i widget al layout
        file_layout.addWidget(self.file_text)
        file_layout.addWidget(self.browse_button)
        left_layout.addLayout(file_layout)
        
        # Riga pulsanti
        buttons_layout = QHBoxLayout()
        self.config_button = QPushButton("Configura")
        self.config_button.clicked.connect(self.on_config_clicked)
        
        self.start_button = QPushButton("Avvia")
        self.start_button.clicked.connect(self.on_start_clicked)
        self.start_button.setEnabled(False)  # Disabilita inizialmente
        
        buttons_layout.addWidget(self.config_button)
        buttons_layout.addWidget(self.start_button)
        buttons_layout.addStretch()
        
        left_layout.addLayout(buttons_layout)
        left_layout.addStretch()
        
        return left_group
    
    def create_right_panel(self):
        """
        Crea il pannello destro con il log.
        
        Returns:
            QGroupBox: Il pannello destro.
        """
        right_group = QGroupBox("Log operazioni")
        right_layout = QVBoxLayout(right_group)
        
        # Widget per il log
        self.log_widget = LogWidget()
        
        # Imposta la policy di dimensionamento
        right_group.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        
        # Aggiungi il widget al layout
        right_layout.addWidget(self.log_widget)
        
        return right_group
    
    def create_button_bar(self):
        """
        Crea la barra dei pulsanti inferiore.
        
        Returns:
            QWidget: Widget contenente i pulsanti.
        """
        button_widget = QWidget()
        button_widget.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Fixed)
        button_layout = QHBoxLayout(button_widget)
        button_layout.setAlignment(Qt.AlignLeft)
        
        # Pulsanti
        self.config_button2 = QPushButton("Configura")
        self.config_button2.setFixedWidth(80)
        self.config_button2.clicked.connect(self.on_config_clicked)
        
        self.reset_button = QPushButton("Reset")
        self.reset_button.setFixedWidth(80)
        self.reset_button.clicked.connect(self.on_reset_clicked)
        
        self.exit_button = QPushButton("Esci")
        self.exit_button.setFixedWidth(80)
        self.exit_button.clicked.connect(self.close)
        
        # Aggiungi al layout
        button_layout.addWidget(self.config_button2)
        button_layout.addWidget(self.reset_button)
        button_layout.addWidget(self.exit_button)
        button_layout.addStretch()
        
        return button_widget
    
    def setup_status_bar(self):
        """Configura la barra di stato."""
        self.statusBar = QStatusBar()
        self.setStatusBar(self.statusBar)
        
        # Aggiunge una progress bar alla status bar
        self.status_progress = QProgressBar()
        self.status_progress.setMaximumWidth(150)
        self.status_progress.setMaximumHeight(16)
        self.status_progress.setValue(0)
        self.statusBar.addPermanentWidget(self.status_progress)
        
        # Stile della barra di stato
        self.statusBar.setStyleSheet("""
            QStatusBar {
                border-top: 1px solid #aaaaaa;
                background-color: #f0f0f0;
            }
            QStatusBar QLabel {
                padding: 0 5px;
                color: #0078d7;
            }
            QStatusBar QProgressBar {
                border: 1px solid #aaaaaa;
                border-radius: 2px;
                background-color: #ffffff;
            }
        """)
        
        self.statusBar.showMessage("Pronto")
    
    def setup_signals(self):
        """Configura i segnali tra i componenti."""
        # Collega il segnale di cambio date
        self.date_widget.dateRangeChanged.connect(self.on_date_changed)
    
    def load_config(self):
        """Carica la configurazione dell'applicazione."""
        self.config = self.config_manager.get_config()
        
        # Assegna lo stato delle operazioni a variabili di classe
        operations = self.config.get("operations", {})
        self.estrai_adm_enabled = operations.get("estrai_AdM", False)
        self.estrai_odm_enabled = operations.get("estrai_OdM", False)
        self.estrai_AFKO_enabled = operations.get("estrai_AFKO", False)
        self.elabora_xls_enabled = operations.get("elabora_xls", False)
    
    def on_log_message(self, message, level):
        """
        Callback per gestire i messaggi di log nell'UI.
        
        Args:
            message (str): Messaggio da visualizzare.
            level (str): Livello del messaggio ('info', 'warning', 'error', ecc.).
        """
        self.log_widget.add_log(message, level)
    
    def update_status_bar(self, message):
        """
        Callback per aggiornare la barra di stato.
        
        Args:
            message (str): Messaggio da visualizzare nella barra di stato.
        """
        self.statusBar.showMessage(message)
    
    def on_date_changed(self, start_date, end_date, is_valid):
        """
        Gestisce il cambio di data nell'intervallo.
        
        Args:
            start_date: Data di inizio.
            end_date: Data di fine.
            is_valid: Se l'intervallo è valido.
        """
        # Abilita o disabilita il pulsante di avvio in base alla validità delle date
        # e alla presenza di un file Excel
        self.update_start_button_state(is_valid)
    
    def update_start_button_state(self, dates_valid=True):
        """
        Aggiorna lo stato del pulsante di avvio.
        
        Args:
            dates_valid (bool): Se le date sono valide.
        """
        # Converte esplicitamente dates_valid in booleano
        dates_valid_bool = bool(dates_valid) if not isinstance(dates_valid, bool) else dates_valid
        
        # Verifica che file_selected sia un booleano
        file_selected = hasattr(self, 'excel_file_path') and self.excel_file_path
        file_selected_bool = bool(file_selected)
        # Abilita il pulsante di avvio solo se entrambe le condizioni sono vere
        if dates_valid_bool and file_selected_bool:
            self.start_button.setEnabled(True)

    def select_excel_file(self):
        """Apre un dialogo per selezionare il file Excel."""
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Seleziona file Excel",
            "",
            "File Excel (*.xlsx *.xls *.xlsm)"
        )
        
        if file_path:
            # Estrai solo il nome del file dal percorso completo
            file_name = os.path.basename(file_path)
            self.log_manager.log(f"Selezionato file excel: {file_name}")
            
            # Mostra il nome del file nel campo di testo
            self.file_text.setText(file_name)
            
            # Salva il percorso completo come attributo dell'oggetto
            self.excel_file_path = file_path
            
            # Verifica lo sheet
            self.log_manager.log("Verifico file excel", "loading", update_status=True, update_log=True)
            
            # Processo il file Excel
            result = self.process_excel_file(file_path, constants.required_sheet, constants.required_columns)
            
            if result:
                # Abilita il pulsante di avvio
                self.update_start_button_state()
                self.log_manager.log("File excel caricato correttamente", "success")
            else:
                self.update_start_button_state(False)
                # Se la verifica fallisce, pulisci il campo
                self.file_text.setText("")
                self.excel_file_path = None
        else:
            self.log_manager.log("Nessun file selezionato", "warning")
    
    def process_excel_file(self, file_path, required_sheet, required_columns):
        """
        Processa il file Excel utilizzando il data processor.
        
        Args:
            file_path (str): Percorso del file Excel.
            required_sheet (str): Nome dello sheet richiesto.
            required_columns (list): Lista delle colonne richieste.
            
        Returns:
            bool: True se il processing è riuscito, False altrimenti.
        """
        return self.excel_data_processor.process_excel_file(file_path, required_sheet, required_columns)
    
    def on_reset_clicked(self):
        """Resetta i campi di input e il log."""
        self.file_text.clear()
        self.date_widget.start_date_picker.setDate(QDate.currentDate())
        self.date_widget.end_date_picker.setDate(QDate.currentDate())
        self.log_widget.clear_logs()
        self.log_manager.log("Eseguito reset dell'applicativo")
        self.start_button.setEnabled(False)
        self.excel_file_path = None  # Resetta il percorso del file Excel
    
    def on_config_clicked(self):
        """Apre la finestra di configurazione."""
        self.statusBar.showMessage("Apertura configurazione...")
        
        config_dialog = ConfigDialog(self.config_manager, self)
        
        # Collega il segnale di log attività al LogManager
        config_dialog.activityLogged.connect(
            lambda message, level: self.log_manager.log(message, level)
        )
        
        if config_dialog.exec_():
            # Ricarica la configurazione se è stata salvata
            self.load_config()
            self.log_manager.log("Configurazione aggiornata")
        else:
            self.log_manager.log("Configurazione non modificata")
    
    def on_start_clicked(self):
        """Avvia l'estrazione dei dati."""
        self.log_manager.log("Avvio estrazioni SAP", origin=logger.name)
        
        # Verifico se è stato selezionato un file
        self.log_manager.log("Verifico selezione file excel")
        if not hasattr(self, 'excel_file_path') or not self.excel_file_path:
            QMessageBox.warning(
                self, 
                "Nessun file selezionato", 
                "Seleziona un file Excel prima di procedere."
            )
            self.log_manager.log("Nessun file excel selezionato", "critical", origin=logger.name)
            return
        
        self.log_manager.log("Verifica file excel - OK", "success", origin=logger.name)
        
        # Verifica la validità delle date inserite
        if not self.date_widget.validate_date_range():
            self.log_manager.log("Errore nella verifica delle date inserite", "critical")
            return
        
        # Ottieni la data di inizio e fine
        start_date, end_date = self.date_widget.get_date_range()
        
        # Ottieni la configurazione delle tecnologie
        tech_config = self.config.get("technologies", {})
        
        # Verifica la configurazione delle tecnologie
        if not self.validate_technology_config(tech_config):
            self.log_manager.log("Errore nella verifica dei codici tecnologia", "critical", origin=logger.name)
            return
        
        # Test stato configurazione
        self.log_manager.log(f"Stato configurazione Check Box AdM: {self.estrai_adm_enabled}", "info", origin=logger.name)
        self.log_manager.log(f"Stato configurazione Check Box OdM: {self.estrai_odm_enabled}", "info", origin=logger.name)
        self.log_manager.log(f"Stato configurazione Check Box AdM: {self.estrai_AFKO_enabled}", "info", origin=logger.name)
        
        # Estraggo i dati da SAP
        self.log_manager.log("Avvio estrazione SAP...", origin=logger.name)
        # Inizializzo i dataframe
        df_IW29 = pd.DataFrame()
        df_IW39 = pd.DataFrame()
        df_AFKO = pd.DataFrame()

        try:
            # Verifica della directory di salvataggio
            save_dir = self.config.get("save_directory", "")
            if not save_dir or not os.path.exists(save_dir):
                self.log_manager.log("Errore: Directory di salvataggio non valida", "critical", origin=logger.name)
                return
            
            # Utilizza la factory per creare una connessione SAP
            with SAPGuiConnection() as sap:
                if sap.is_connected():
                    session = sap.get_session()
                    if session:
                        self.log_manager.log("Connessione SAP attiva")
                        
                        # Crea un estrattore di dati SAP
                        extractor = SAPDataExtractor(session, self)
                        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                        
                        # Verifico la check box per l'estrazione AdM self.estrai_adm_enabled
                        if self.estrai_adm_enabled:
                            # Estrazione dati AdM
                            self.log_manager.log("Estrazione dati IW29", "info")
                            result, df_IW29 = extractor.extract_IW29(
                                start_date, 
                                end_date, 
                                tech_config, 
                                self.excel_data_processor.get_adm()["idItem"]
                            )
                            
                            if not result:
                                self.log_manager.log("Errore: Estrazione IW29 fallita", "error")
                                return
                            
                            # Salva il DataFrame in un file Excel
                            output_file = os.path.join(save_dir, f"IW29_AdM_{timestamp}.xlsx")
                            try:
                                df_IW29.to_excel(output_file, index=False)
                                self.log_manager.log(f"File salvato in: {output_file}", "success")
                            except Exception as e:
                                self.log_manager.log(f"Errore: Salvataggio file IW29 fallito: {str(e)}", "error")
                                return
                        
                        # Verifico la check box per l'estrazione OdM self.estrai_odm_enabled
                        if self.estrai_odm_enabled:
                            # Estrazione dati OdM
                            self.log_manager.log("Estrazione dati IW39")
                            result, df_IW39 = extractor.extract_IW39(
                                start_date, 
                                end_date, 
                                tech_config, 
                                self.excel_data_processor.get_odm()
                            )
                            
                            if not result:
                                self.log_manager.log("Errore: Estrazione IW39 fallita", "error")
                                return
                            
                            # Salva il DataFrame in un file Excel
                            output_file = os.path.join(save_dir, f"IW39_OdM_{timestamp}.xlsx")
                            try:
                                df_IW39.to_excel(output_file, index=False)
                                self.log_manager.log(f"File salvato in: {output_file}", "success")
                            except Exception as e:
                                self.log_manager.log(f"Errore: Salvataggio file IW39 fallito: {str(e)}", "error")
                                return
                            
                        # Verifico la check box self.estrai_AFKO_enabled per l'estrazione delle date inizio cardine degli OdM relativi al DF AdM     
                        if (self.estrai_AFKO_enabled):
                            if (df_IW29 is not None and not df_IW29.empty):
                                # Estrazione dati dal dataframe df_IW29
                                df_OdM = pd
                                df_AdM = (df_IW29["Avvisi"])
                                self.log_manager.log("Estrazione dati SE16 tabella AFKO", "info")
                                # Ricavo la lista degli OdM presenti nel 
                                result, df_AFKO = extractor.extract_SE16(df_AdM)
                                
                                if not result:
                                    self.log_manager.log("Errore: Estrazione AFKO fallita", "error")
                                    return
                                
                                # Salva il DataFrame in un file Excel
                                output_file = os.path.join(save_dir, f"AFKO_{timestamp}.xlsx")
                                try:
                                    df_AFKO.to_excel(output_file, index=False)
                                    self.log_manager.log(f"File salvato in: {output_file}", "success")
                                except Exception as e:
                                    self.log_manager.log(f"Errore: Salvataggio file AFKO fallito: {str(e)}", "error")
                                    return
                            else:
                                self.log_manager.log("Errore estrazione dati SE16 - tabella df_IW29 non esistente.", "errore")
                                return

                        # Estrazione dati completata
                        self.log_manager.log("Estrazione completata con successo", "success")
                        # Elaboro i dati estratti
                        # Verifico i dati ottenuti
                        if ((df_IW29 is not None and not df_IW29.empty) and
                            (df_IW39 is not None and not df_IW39.empty) and
                            (df_AFKO is not None and not df_AFKO.empty)):
                            self.log_manager.log("Elaboro i dati estratti", "info")
                            # Elaboro i dati
                            result, data = self.process_data(df_IW29, df_IW39, df_AFKO)
                            if result:
                                self.log_manager.log("Elaborazione dati completata con successo", "success")
                        else:
                            self.log_manager.log("Errore: Dati estratti vuoti o non validi", "error")
                            return

                else:
                    self.log_manager.log("Connessione SAP NON attiva", "error")
                    return
        except Exception as e:
            self.log_manager.log(f"Estrazione dati SAP: Errore: {str(e)}", "error")
            return
    
    def process_data(self, df_IW29, df_IW39, df_AFKO) -> Tuple[bool, Dict[str, pd.DataFrame]]:
        """
        Elabora i dati estratti da SAP.
        
        Args:
            df_IW29 (DataFrame): DataFrame contenente i dati IW29.
            df_IW39 (DataFrame): DataFrame contenente i dati IW39.
            df_AFKO (DataFrame): DataFrame contenente i dati AFKO.
            
        Returns:
            Tuple[bool, Dict[str, pd.DataFrame]]: 
            - bool: True se almeno un DataFrame è stato estratto con successo
            - Dict: Dizionario con chiavi come nomi delle tabelle e valori come DataFrame
        """
        self.log_manager.log("Inizio elaborazione dati estratti", "info")
        # Inizializza i DataFrame
        dataframes = {
            'AdM': pd.DataFrame(),
            'OdM': pd.DataFrame(), 
            'IW29': pd.DataFrame(),
            'IW39': pd.DataFrame()
        }        
        # Esegui l'elaborazione dei dati
        # Inserisco una nuova colonna "Data inizio cardine" nel df_IW29, ricavo il data di inizio cardine dal df_AFKO
        df_IW29["Data_inizio_cardine"] = df_IW29["Avvisi"].map(df_AFKO.set_index("Avvisi")["Data inizio cardine"])    
        return True, None

    def validate_technology_config(self, tech_config):
        """
        Verifica che le tecnologie siano configurate correttamente.
        
        Args:
            tech_config (dict): Configurazione delle tecnologie.
            
        Returns:
            bool: True se la configurazione è valida, False altrimenti.
        """
        self.log_manager.log(f"Verifica tecnologie configurate")
        
        # Verifica che tech_config non sia vuoto
        if not tech_config:
            error_message = "Nessuna tecnologia configurata. Utilizzare la finestra di configurazione."
            QMessageBox.warning(self, "Configurazione mancante", error_message)
            self.log_manager.log("Nessuna tecnologia configurata", "error")
            return False
        
        # Controlla se ci sono tecnologie con prefissi configurati
        has_valid_tech = False
        for tech, prefixes in tech_config.items():
            if prefixes:  # Se c'è almeno un prefisso configurato
                has_valid_tech = True
                break
        
        if not has_valid_tech:
            error_message = "Tutte le tecnologie sono prive di prefissi. Configurare almeno un prefisso."
            QMessageBox.warning(self, "Configurazione incompleta", error_message)
            self.log_manager.log(error_message, "error")
            return False
        
        # Se arriviamo qui, la configurazione è valida
        self.log_manager.log("Verifica tecnologie configurate - OK", "success")
        return True
    
    def closeEvent(self, event):
        """
        Gestisce l'evento di chiusura della finestra.
        
        Args:
            event: Evento di chiusura.
        """
        self.log_manager.log("Applicazione terminata", "info")
        event.accept()