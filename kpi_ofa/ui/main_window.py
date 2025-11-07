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
from kpi_ofa.core.class_excel_file_mng import ExcelFileMng
from kpi_ofa.core.excel_data_processor import ExcelDataProcessor
from kpi_ofa.core.class_df_processor import DfProcessor
from kpi_ofa.core.log_manager import LogManager
from kpi_ofa.core.test_data_loader import TestDataLoader, TestDataLoadError
from kpi_ofa.services.sap_connection import SAPGuiConnection
from kpi_ofa.services.sap_transactions import SAPDataExtractor

import kpi_ofa.config.constants as constants

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

        # === DATAFRAMES PRINCIPALI ===
        # Inizializza tutti i DataFrame come vuoti
        self._init_dataframes()        

        # Configura la finestra
        self.setup_window()    

        # Crea l'interfaccia utente
        self.setup_ui()
        
        # Inizializza i componenti di supporto
        self.init_components()
        
        # Configura i segnali
        self.setup_signals()

        # Inizializza TestDataLoader con dependency injection
        self.test_data_loader = TestDataLoader(main_window=self)
        self.dataframes = {} # Dizionario per memorizzare i DataFrame elaborati
        
        # Carica la configurazione
        self.load_config()

        # Elabora file Excel
        self.excel_file_manager = ExcelFileMng(self.data_directory)    

        # Inizializza la modalità di debug
        self._init_debug_mode()

    def _init_debug_mode(self):
        """
        Inizializza tutti i DataFrame dell'applicazione.
        
        Questo metodo centralizza l'inizializzazione di tutti i DataFrame
        per una gestione più pulita e organizzata.
        """
        # Carica i file di test se in modalità debug
        if constants.DEBUG_MODE:
            try:
                # Carica tutti i file di test
                success = self.test_data_loader.load_test_files()
                
                if success:
                    self.log_manager.log("🎯 File di test caricati con successo", "success", origin=logger.name)
                    self.disable_buttons()  # Disabilita i pulsanti in modalità debug
                    # Messaggio di avvio
                    self.log_manager.log("Sistema avviato in modalità DEBUG!", "info", origin=logger.name)
                    # Ora puoi usare i DataFrame direttamente
                    if hasattr(self.test_data_loader, 'df_IW29'):
                        rows = len(self.test_data_loader.df_IW29)
                        self.log_manager.log(f"📊 IW29: {rows} righe disponibili", "info", origin=logger.name)
                    
            except TestDataLoadError as e:
                # Errore specifico di caricamento
                self.log_manager.log(f"❌ Errore caricamento file di test: {str(e)}", "error", origin=logger.name)
                self.start_button.setEnabled(False)
                return
                
            except Exception as e:
                # Altri errori inaspettati
                self.log_manager.log(f"💥 Errore critico: {str(e)}", "error", origin=logger.name)
                self.start_button.setEnabled(False)
                return
        else:
            self.start_button.setEnabled(False)
            # Messaggio di avvio
            self.log_manager.log("Sistema pronto!", "info", origin=logger.name)

    def _init_dataframes(self):
        """
        Inizializza tutti i DataFrame dell'applicazione.
        
        Questo metodo centralizza l'inizializzazione di tutti i DataFrame
        per una gestione più pulita e organizzata.
        """
        # === DATAFRAMES DA SAP ===
        self.df_IW29 = pd.DataFrame()           # Avvisi di manutenzione ottenuti con # transazione IW29
        self.df_IW39 = pd.DataFrame()           # Ordini di manutenzione ottenuti con # transazione IW39
        self.df_AFKO = pd.DataFrame()           # Date inizio cardine ottenuta con transazione S16 tabella AFKO
        
        # === DATAFRAMES DA FILE EXCEL ===
        self.df_excel_normalized = pd.DataFrame()  # Excel normalizzato
        self.df_plants = pd.DataFrame()           # Connettività impianti
        
        # === DATAFRAMES ELABORATI ===
        self.df_merged = pd.DataFrame()         # Dati uniti/elaborati
        self.df_final_report = pd.DataFrame()  # Report finale
        
        # === DATAFRAMES AUSILIARI ===
        self.df_AdM = pd.DataFrame()            # Avvisi di manutenzione filtrati da file excel normalizzato
        self.df_OdM = pd.DataFrame()            # Ordini di manutenzione filtrati da file excel normalizzato
        
        self.log_manager.log("DataFrame inizializzati come vuoti", "info", origin=logger.name)
    
    def init_components(self):
        """Inizializza i componenti di supporto dell'applicazione."""
        # Configura i callback per il LogManager
        self.log_manager.set_ui_callback(self.on_log_message)
        self.log_manager.set_status_bar_callback(self.update_status_bar)
        
        # Processore di dati
        self.excel_data_processor = ExcelDataProcessor()    
        
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
        # Salvo i percorsi delle directory leggendole dal file di configurazione
        self.data_directory = self.config.get("data_directory", "")
        self.sap_directory = os.path.join(self.data_directory, "SAP")
        self.powerbi_directory = os.path.join(self.data_directory, "PowerBI")
        self.test_directory = os.path.join(self.data_directory, "test")    

        # Aggiorna configurazione in TestDataLoader
        if hasattr(self, 'test_data_loader'):
            self.test_data_loader.set_config(self.config)        
        
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

    def disable_buttons(self):
        """Disabilita i pulsanti di avvio e configurazione nella modalità di debug"""
        self.start_button.setEnabled(False)
        self.config_button.setEnabled(False)
        self.config_button2.setEnabled(False)
        self.browse_button.setEnabled(False)
        self.date_widget.start_date_picker.setEnabled(False)
        self.date_widget.end_date_picker.setEnabled(False)
        self.start_button.setEnabled(True)

    def enable_buttons(self):
        """Abilita i pulsanti di avvio e configurazione nella modalità di debug"""
        self.start_button.setEnabled(True)
        self.config_button.setEnabled(True)
        self.config_button2.setEnabled(True)
        self.browse_button.setEnabled(True)
        self.date_widget.start_date_picker.setEnabled(True)
        self.date_widget.end_date_picker.setEnabled(True)
        self.update_start_button_state(False)    


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
            result, self.df_excel_normalized = self.process_excel_file(file_path, constants.required_sheet, constants.required_columns)
            if (result == False):
                self.log_manager.log("Errore durante l'elaborazione del file excel", "error")
                self.update_start_button_state(False)
                # Se la verifica fallisce, pulisci il campo
                self.file_text.setText("")
                self.excel_file_path = None
                return
            else:
                # Ottieni la data di inizio e fine dall'intervallo di date
                start_date, end_date = self.date_widget.get_date_range()
                # Aggiungi la colonna 'Data_fine_estrazione' al df
                self.df_excel_normalized['Data_fine_estrazione'] = end_date.toString("dd.MM.yyyy")  # Formato gg.mm.aaaa
                # Log di successo
                self.log_manager.log("File excel verificato correttamente", "success")
                self.update_start_button_state()
                # Salvo il df in un file excel per eventuali elaborazioni successive
                try:
                    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                    # Salva il DataFrame in un file Excel
                    output_file = os.path.join(self.data_directory, f"df_excel_norm_{timestamp}.xlsx")
                    #self.df_excel_normalized.to_excel(output_file, index=False)
                    if not self.excel_file_manager.save_excel_file_advanced(self.df_excel_normalized, output_file):
                        raise Exception("Salvataggio file fallito")
                    self.log_manager.log(f"File salvato in: {output_file}", "success")
                except Exception as e:
                    self.log_manager.log(f"Errore: Salvataggio file {output_file} fallito: {str(e)}", "error")
                    return    
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
        
    # def load_excel_file(self, file_path, sheet_name=0):
    #     """Carica un file Excel con gestione errori"""
    #     try:
    #         # Verifica che il file esista
    #         if not os.path.exists(file_path):
    #             self.log_manager.log(f"❌ File non trovato: {file_path}")
    #             return None
            
    #         # Carica il file
    #         df = pd.read_excel(file_path, sheet_name=sheet_name)
    #         self.log_manager.log(f"✅ File caricato: {len(df)} righe, {len(df.columns)} colonne")
    #         return df
            
    #     except FileNotFoundError:
    #         self.log_manager.log(f"❌ File non trovato: {file_path}")
    #         return None
    #     except PermissionError:
    #         self.log_manager.log(f"❌ Permessi insufficienti per leggere: {file_path}")
    #         return None
    #     except Exception as e:
    #         self.log_manager.log(f"❌ Errore nel caricamento: {str(e)}")
    #         return None
    
    def on_reset_clicked(self):
        """Resetta i campi di input e il log."""
        self.file_text.clear()
        self.date_widget.start_date_picker.setDate(QDate.currentDate())
        self.date_widget.end_date_picker.setDate(QDate.currentDate())
        self.log_widget.clear_logs()
        self.log_manager.log("Eseguito reset dell'applicativo")
        self.excel_file_path = None  # Resetta il percorso del file Excel

        # Carico i file di test se in modalità debug
        self._init_debug_mode()	
    
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
        # Verifica se il sistema è in modalità debug
        if constants.DEBUG_MODE:
            """Elabora dati solo se tutti i file sono caricati."""
            if not self.test_data_loader.is_loaded():
                self.log_warning("❌ Non tutti i file di test sono caricati")
                return False
            
            # Se arriviamo qui, TUTTI i DataFrame sono disponibili
            self.df_IW29 = self.test_data_loader.df_IW29
            self.df_IW39 = self.test_data_loader.df_IW39
            self.df_AFKO = self.test_data_loader.df_AFKO
            self.df_excel_normalized = self.test_data_loader.df_excel_normalized
            self.df_plants = self.test_data_loader.df_plants

        else:
            # se non siamo in modalità debug, procediamo con l'estrazione dei dati da SAP
            """Avvia l'estrazione dei dati."""
            self.log_manager.log("Avvio estrazioni SAP", origin=logger.name)
            # Misura il tempo
            start_time = time.perf_counter()

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
            self.log_manager.log(f"Stato configurazione Check Box AFKO: {self.estrai_AFKO_enabled}", "info", origin=logger.name)
            
            # Estraggo i dati da SAP
            self.log_manager.log("Avvio estrazione SAP...", origin=logger.name)
            try:
                
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
                                result, self.df_IW29 = extractor.extract_IW29(
                                    start_date, 
                                    end_date, 
                                    tech_config, 
                                    self.excel_data_processor.get_adm()["AdM"]
                                )
                                
                                if not result:
                                    self.log_manager.log("Errore: Estrazione IW29 fallita", "error")
                                    return
                                
                                # Aggiungo la colonna 'Data_fine_estrazione' con la data di fine estrazione per costurire il report in PowerBI
                                self.df_IW29['Data_fine_estrazione'] = end_date.toString("dd.MM.yyyy")  # Formato gg.mm.aaaa
                                
                                # Salva il DataFrame in un file Excel
                                output_file = os.path.join(self.sap_directory, f"IW29_AdM_{timestamp}.xlsx")
                                try:
                                    #self.df_IW29.to_excel(output_file, index=False)
                                    if not self.excel_file_manager.save_excel_file_advanced(self.df_IW29, output_file):
                                        raise Exception("Salvataggio file fallito")
                                    self.log_manager.log(f"File salvato in: {output_file}", "success")
                                except Exception as e:
                                    self.log_manager.log(f"Errore: Salvataggio file IW29 fallito: {str(e)}", "error")
                                    return
                            
                            # Verifico la check box per l'estrazione OdM self.estrai_odm_enabled
                            if self.estrai_odm_enabled:
                                # Estrazione dati OdM
                                self.log_manager.log("Estrazione dati IW39")
                                result, self.df_IW39 = extractor.extract_IW39(
                                    start_date, 
                                    end_date, 
                                    tech_config, 
                                    self.excel_data_processor.get_odm()["OdM"]
                                )
                                
                                if not result:
                                    self.log_manager.log("Errore: Estrazione IW39 fallita", "error")
                                    return

                                # Aggiungo la colonna 'Data_fine_estrazione' con la data di fine estrazione per costurire il report in PowerBI
                                self.df_IW39['Data_fine_estrazione'] = end_date.toString("dd.MM.yyyy")   # Formato gg.mm.aaaa
                                
                                # Salva il DataFrame in un file Excel
                                output_file = os.path.join(self.sap_directory, f"IW39_OdM_{timestamp}.xlsx")
                                try:
                                    #self.df_IW39.to_excel(output_file, index=False)
                                    if not self.excel_file_manager.save_excel_file_advanced(self.df_IW39, output_file):
                                        raise Exception("Salvataggio file fallito")                                    
                                    self.log_manager.log(f"File salvato in: {output_file}", "success")
                                except Exception as e:
                                    self.log_manager.log(f"Errore: Salvataggio file IW39 fallito: {str(e)}", "error")
                                    return
                                
                            # Verifico la check box self.estrai_AFKO_enabled per l'estrazione delle date inizio cardine degli OdM relativi al DF AdM     
                            if (self.estrai_AFKO_enabled):
                                # TEST 
                                # Se non è stata compiuta l'estrazione AdM, allora considero il file Excel di una estrazione precedente.
                                if not(self.estrai_adm_enabled):
                                    self.log_manager.log("Estrazione AdM non effettuata. Impossibile estrarre AFKO.", "error")
                                    return False
                                
                                # Procedo con l'estrazione degli OdM dal dataframe creato con l'estrazione IW29
                                if (self.df_IW29 is not None and not self.df_IW29.empty):
                                    try:
                                        # Estrazione dati dal dataframe df_IW29
                                        df_OdM = (self.df_IW29["Ordine"]          
                                            .astype(str)                           # Converte tutto in stringa
                                            .str.strip()                           # Rimuove spazi
                                            .replace(['', 'nan', 'NaN', 'None', 'null'], pd.NA)  # Sostituisce valori vuoti
                                            .pipe(pd.to_numeric, errors='raise')  # Conversione sicura (NaN per errori)
                                            .dropna()                              # Rimuove NaN dalla conversione
                                            .astype('Int64'))                      # Converte in Int64
                                    except Exception as e:
                                        self.log_manager.log(f"Errore durante la conversione degli Ordini in df_IW29: {str(e)}", "error")
                                        return 

                                # Numero di OdM estratti
                                num_odm = df_OdM.drop_duplicates().shape[0]
                                self.log_manager.log(f"Totale OdM = {num_odm}", "info")
                                self.log_manager.log("Estrazione dati SE16 tabella AFKO", "info")
                                # Ricavo la lista degli OdM presenti nel
                                result, self.df_AFKO = extractor.extract_SE16(df_OdM)
                                
                                if not result:
                                    self.log_manager.log("Errore: Estrazione AFKO fallita", "error")
                                    return
                                
                                # Aggiungo la colonna 'Data_fine_estrazione' con la data di fine estrazione per costurire il report in PowerBI
                                self.df_AFKO['Data_fine_estrazione'] = end_date.toString("dd.MM.yyyy")   # Formato gg.mm.aaaa                             
                                
                                # Verifico che il numero di OdM estratti sia uguale a quello di OdM presenti nel df_IW29
                                if self.df_AFKO is not None and not self.df_AFKO.empty:
                                    num_afko = self.df_AFKO.shape[0]
                                    if num_afko != num_odm:
                                        self.log_manager.log(f"Attenzione: Numero di OdM estratti ({num_afko}) diverso da quello atteso ({num_odm})", "error")
                                        return
                                    else:
                                        self.log_manager.log(f"Numero di OdM estratti ({num_afko}) corrisponde a quello atteso ({num_odm})", "success")
                                        # Salva il DataFrame in un file Excel
                                        output_file = os.path.join(self.sap_directory, f"df_AFKO_{timestamp}.xlsx")
                                        try:
                                            #self.df_AFKO.to_excel(output_file, index=False)
                                            if not self.excel_file_manager.save_excel_file_advanced(self.df_AFKO, output_file):
                                                raise Exception("Salvataggio file fallito")                                            
                                            self.log_manager.log(f"File salvato in: {output_file}", "success")
                                        except Exception as e:
                                            self.log_manager.log(f"Errore: Salvataggio file AFKO fallito: {str(e)}", "error")
                                            return                                    
                                else:
                                    self.log_manager.log("Errore estrazione dati SE16 - tabella df_IW29 non esistente.", "errore")
                                    return

                            # Estrazione dati completata
                            self.log_manager.log("Estrazione completata con successo", "success")
                            # Calcola il tempo di esecuzione
                            end_time = time.perf_counter()
                            execution_time = end_time - start_time
                            time_str = self.format_execution_time(execution_time)
                            self.log_manager.log(f"Tempo estrazioni SAP: {time_str}", "info")
                            #return
                    else:
                        self.log_manager.log("Connessione SAP NON attiva", "error")
                        return
            except Exception as e:
                self.log_manager.log(f"Estrazione dati SAP: Errore: {str(e)}", "error")
                return
            
            # Elaboro i dati estratti
            self.log_manager.log("Avvio elaborazione dati...", origin=logger.name)
            try:
                # Carico il file excel plants.xlsx in un df
                file_path = os.path.join(self.data_directory, 'plants.xlsx')
                result, self.df_plants = self.excel_file_manager.load_excel_file(file_path)
                if not result:
                    self.log_manager.log(f"Errore durante il caricamento del file {file_path}", "error")
                    return
                # Creo un dizionario contenente tutti i df necessari all'elaborazione.
                dict_df = {
                    "df_AFKO": self.df_AFKO,
                    "df_excel_normalized": self.df_excel_normalized,
                    "df_IW29": self.df_IW29,
                    "df_IW39": self.df_IW39,
                    "df_plants": self.df_plants
                }

                # Salvo le estrazioni nella directory SAP 
                exclude_list = ["df_excel_normalized", "df_plants"]
                filtered_dict = {k: v for k, v in dict_df.items() if k not in exclude_list}
                if not self.save_all_dataframes(filtered_dict, self.sap_directory, "SAP"):
                    self.log_manager.log("Errore durante il salvataggio delle estrazioni nella directory SAP", "error")
                    return

                # Procedo con l'elaborazione dei dati
                result, data = self.process_data(dict_df)
                if not result:
                    self.log_manager.log("Errore durante l'elaborazione dei dati estratti", "error")
                    return
                self.log_manager.log("Elaborazione dati completata con successo", "success")
            except Exception as e:
                self.log_manager.log(f"Errore durante l'elaborazione dei dati: {str(e)}", "error")
                return


    def format_execution_time(self, seconds: float) -> str:
        """
        Formatta il tempo di esecuzione in modo leggibile.
        
        Args:
            seconds: Tempo in secondi
            
        Returns:
            Stringa formattata (es. "2h 30m 15s", "5m 23s", "1.23s", "123ms", "12.3μs")
        """
        if seconds >= 3600:  # >= 1 ora
            hours = int(seconds // 3600)
            minutes = int((seconds % 3600) // 60)
            remaining_seconds = seconds % 60
            
            if minutes > 0 and remaining_seconds >= 1:
                return f"{hours}h {minutes}m {remaining_seconds:.0f}s"
            elif minutes > 0:
                return f"{hours}h {minutes}m"
            else:
                return f"{hours}h {remaining_seconds:.1f}s"
                
        elif seconds >= 60:  # >= 1 minuto
            minutes = int(seconds // 60)
            remaining_seconds = seconds % 60
            
            if remaining_seconds >= 1:
                return f"{minutes}m {remaining_seconds:.0f}s"
            else:
                return f"{minutes}m {remaining_seconds:.1f}s"
                
        elif seconds >= 1.0:  # >= 1 secondo
            return f"{seconds:.2f}s"
        elif seconds >= 0.001:  # >= 1 millisecondo
            return f"{seconds * 1000:.1f}ms"
        elif seconds >= 0.000001:  # >= 1 microsecondo
            return f"{seconds * 1000000:.1f}μs"
        else:  # < 1 microsecondo
            return f"{seconds * 1000000000:.0f}ns"

    def process_data(self, dict_df) -> Tuple[bool, Dict[str, pd.DataFrame]|None]:
        """
        Elabora i dati estratti da SAP.
        
        Args:
            df_dict: Dizionario) contenente i DataFrame:
                df_IW29: DataFrame contenente i dati IW29.
                df_IW39: DataFrame contenente i dati IW39.
                df_AFKO: DataFrame contenente i dati AFKO.
                df_excel_normalized: Dataframe contenente i dati normalizzato dal file Excel di input.
                df_plants: Dataframe contenente i dati relativi ai plants.    
            
        Returns:
            Tuple[bool, Dict[str, pd.DataFrame]]: 
            - bool: True se almeno un DataFrame è stato estratto con successo
            - Dict: Dizionario con chiavi come nomi delle tabelle e valori come DataFrame
        """
        # Ottieni la data di inizio e fine dall'intervallo di date
        intervallo_date = self.date_widget.get_date_range()
        if len(intervallo_date) >= 2:
            data_inizio = intervallo_date[0].toString("dd.MM.yyyy")
            data_fine = intervallo_date[1].toString("dd.MM.yyyy")
        else:
            print("Errore: intervallo date non valido")
            return False, None

        # creo una istanza di DfProcessor per elaborare i DataFrame
        df_processor = DfProcessor(dict_df, intervallo_date)
        df_processor.get_dataframe_info()


        # Converti tutti i campi dei df
        print("\nInizio conversione dei DataFrame...")
        result, dict_df_conv = df_processor.converti_tutti_df(dict_df)
        if not result:
            print("Errore durante la conversione dei DataFrame")
            return False, None
        print("\n\tConversione completata con successo")
        
        # Inserisco delle nuove colonne nei df per poter effettuare successivamente l'elaborazione
        print("\nInizio inserimento nuove colonne nei DataFrame...")
        result, dict_df_add_columns = df_processor.df_add_columns(dict_df_conv)
        if not result:
            print("Errore durante l'inserimento delle nuove colonne nei DataFrame")
            return False, None
        print("\n\tInserimento completato con successo")
        
        # Salvo i df creati in file excel
        print("\nInizio salvataggio DataFrame in file Excel...")
        if not self.save_all_dataframes(dict_df_add_columns, self.data_directory):
            print("Errore durante il salvataggio dei DataFrame in file Excel")
            return False, None
        else:
            print("\n\tSalvataggio completato con successo")
        
        # Eseguo l'elaborazione dei dati
        print("\nInizio elaborazione dati...")
        result, dict_df_processed = df_processor.process_dataframes(dict_df_add_columns)
        if not result:
            print("Errore durante l'elaborazione dei dati")
            return False, None
        print("\n\tElaborazione completata con successo")

        # Salvo i df elaborati in file excel per utilizzarli nella PowerBI
        print("\nInizio salvataggio DataFrame elaborati in file Excel...")
        if not self.save_all_dataframes(dict_df_processed, self.powerbi_directory, "PowerBI", data_fine):
            print("Errore durante il salvataggio dei DataFrame elaborati in file Excel")
            return False, None

        print("\n\tSalvataggio completato con successo")

        # elaborazione completata con successo
        return True, None        
    
    def save_all_dataframes(self, dict_df, file_path, infix="extended", postfix=""):
        """
        Salva tutti i DataFrame contenuti nel dizionario in file Excel separati.
        
        Args:
            dict_df (dict): Dizionario contenente i DataFrame da salvare
            
        Returns:
            bool: True se tutti i file sono stati salvati con successo, False altrimenti
        """
        print("\nInizio salvataggio DataFrame in file Excel...")
        
        if postfix == "":
            postfix = datetime.now().strftime("%Y%m%d_%H%M%S")
        
        # Contatori per statistiche
        saved_count = 0
        failed_count = 0
        failed_files = []
        
        try:
            # Ciclo attraverso tutti i DataFrame nel dizionario
            for df_name, dataframe in dict_df.items():
                try:
                    # Verifica che il valore sia effettivamente un DataFrame
                    if not isinstance(dataframe, pd.DataFrame):
                        self.log_manager.log(f"Skipping {df_name}: non è un DataFrame", "warning")
                        continue
                    
                    # Verifica che il DataFrame non sia vuoto
                    if dataframe.empty:
                        self.log_manager.log(f"Skipping {df_name}: DataFrame vuoto", "warning")
                        continue
                    
                    # Genera il nome del file output
                    output_file = os.path.join(file_path, f"{df_name}_{infix}_{postfix}.xlsx")
                    
                    # Salva il DataFrame usando il file manager
                    if self.excel_file_manager.save_excel_file_advanced(dataframe, output_file):
                        saved_count += 1
                        self.log_manager.log(f"✓ {df_name} salvato in: {output_file}", "success")
                    else:
                        failed_count += 1
                        failed_files.append(df_name)
                        self.log_manager.log(f"✗ Errore salvando {df_name}", "error")
                        
                except Exception as e:
                    failed_count += 1
                    failed_files.append(df_name)
                    self.log_manager.log(f"✗ Errore salvando {df_name}: {str(e)}", "error")
            
            # Log del riepilogo finale
            total_dataframes = len(dict_df)
            self.log_manager.log(f"Salvataggio completato: {saved_count}/{total_dataframes} file salvati", 
                                "success" if failed_count == 0 else "warning")
            
            if failed_files:
                self.log_manager.log(f"File non salvati: {', '.join(failed_files)}", "error")

        except Exception as e:
            self.log_manager.log(f"Errore generale durante il salvataggio: {str(e)}", "error")
            return False

        # Ritorna True solo se tutti i file sono stati salvati con successo
        return failed_count == 0    

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