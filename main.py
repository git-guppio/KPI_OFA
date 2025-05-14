import sys
import os
import json
import pandas as pd
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                            QHBoxLayout, QLabel, QDateEdit, QFileDialog, QLineEdit,
                            QSizePolicy, QPushButton, QStatusBar, QMessageBox,
                            QListWidget, QGroupBox, QProgressBar, QMenu, QAction,
                            QListWidgetItem, QStyle)
from PyQt5.QtGui import QCursor
from PyQt5.QtCore import QDate, Qt
import SAP_Connection
import SAP_Transactions
import Config.constants as constants
import openpyxl
import logging
import time

# Importazione della classe di configurazione
from config_dialog import ConfigDialog

# Configurazione base del logging per tutta l'applicazione
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler("app.log"),
        logging.StreamHandler()
    ]
)

# Logger specifico per questo modulo
logger = logging.getLogger("main")
logger.setLevel(logging.DEBUG)

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        # Imposta il nome del file di configurazione all'inizio del metodo __init__
        self.config_file = constants.configuration_json

        # Dataframe contenenti i dati dal file Excel attraverso la finestra di selezione file
        self.df_AdM = None
        self.df_OdM = None
        self.excel_file_path = None  # Percorso del file Excel selezionato

        # Imposta i flag della finestra per mostrare solo il pulsante di chiusura
        #self.setWindowFlags(self.windowFlags() & ~Qt.WindowMaximizeButtonHint & ~Qt.WindowMinimizeButtonHint)
            
        self.setWindowTitle("Estrai Dati KPI OFA")
        self.setWindowIconText("KPI OFA")
        self.setGeometry(100, 100, 180, 168)  # x, y, width, height
        
        # Widget centrale
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        # Layout principale
        main_layout = QVBoxLayout(central_widget)

        # Layout orizzontale per i due pannelli
        content_layout = QHBoxLayout()

        # Crea un widget contenitore per il pannello sinistro
        left_group = QGroupBox("Estrazioni SAP")  # Il testo è il titolo del gruppo
        left_layout = QVBoxLayout(left_group)  # Assegna il layout direttamente al gruppo
        # Imposta la policy di dimensionamento per NON permettere l'espansione
        left_group.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
        # Imposta una dimensione fissa (regola in base alle tue esigenze)
        left_group.setFixedWidth(180)
        # Opzione 1: Imposta anche l'altezza fissa se sai esattamente quanto spazio serve
        left_group.setFixedHeight(190)

        # Crea un gruppo contenitore per il pannello destro
        right_group = QGroupBox("Log operazioni")
        right_layout = QVBoxLayout(right_group)
        # Imposta la policy di dimensionamento per permettere l'espansione
        right_group.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
        right_group.setAlignment(Qt.AlignTop)  # Allinea il pannello a sinistra

        # Riga 1 - Data inizio
        row1_layout = QHBoxLayout()
        start_date_label = QLabel("Data inizio:")
        start_date_label.setFixedWidth(60)
        self.start_date_picker = QDateEdit()
        self.start_date_picker.setDate(QDate.currentDate())
        self.start_date_picker.setCalendarPopup(True)
        self.start_date_picker.setMinimumWidth(100)
        row1_layout.addWidget(start_date_label)
        row1_layout.addWidget(self.start_date_picker)
        row1_layout.addStretch()
        
        # Riga 2 - Data fine
        row2_layout = QHBoxLayout()
        end_date_label = QLabel("Data fine:")
        end_date_label.setFixedWidth(60)
        self.end_date_picker = QDateEdit()
        self.end_date_picker.setDate(QDate.currentDate())
        self.end_date_picker.setCalendarPopup(True)
        self.end_date_picker.setMinimumWidth(100)
        row2_layout.addWidget(end_date_label)
        row2_layout.addWidget(self.end_date_picker)
        row2_layout.addStretch()

        # Riga 3 - Selezione file Excel (sostituisce la finestra di testo)
        file_layout = QVBoxLayout()

        # Campo di testo di sola lettura per mostrare il nome del file
        self.file_text = QLineEdit()
        self.file_text.setReadOnly(True)  # Rende il campo non modificabile
        self.file_text.setFixedWidth(165)
        self.file_text.setPlaceholderText("Nessun file selezionato...")

        # Pulsante per selezionare il file
        self.browse_button = QPushButton("Seleziona Excel...")
        self.browse_button.setFixedWidth(165)
        self.browse_button.clicked.connect(self.select_excel_file)

        # Aggiungi i widget al layout
        file_layout.addWidget(self.file_text)  # Proporzione 3
        file_layout.addWidget(self.browse_button)  # Proporzione 1
   
        # Riga 4 - Bottoni
        row4_layout = QHBoxLayout()
        self.config_button = QPushButton("Configura")
        self.config_button.clicked.connect(self.on_config_clicked)
        self.config_button.setFixedWidth(80)        
        self.start_button = QPushButton("Avvia")
        self.start_button.clicked.connect(self.on_start_clicked)
        self.start_button.setFixedWidth(80)
        self.start_button.setEnabled(False)  # Disabilita il pulsante se lo sheet non esiste
        row4_layout.addWidget(self.config_button)
        row4_layout.addWidget(self.start_button)
        row4_layout.addStretch()

        # Aggiungi i layout alla colonna sinistra
        left_layout.addLayout(row1_layout)
        left_layout.addLayout(row2_layout)
        left_layout.addLayout(file_layout)
        left_layout.addLayout(row4_layout)
        # Aggiungi il layout sinistro al layout orizzontale
        content_layout.addWidget(left_group, 0)

        # Aggiungi un'area di testo per il log delle operazioni
        self.log_list = QListWidget()
        self.log_list.setMinimumWidth(200)
        # Imposta la policy di dimensionamento per il campo di testo
        self.log_list.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)        
        # Aggiungi il campo di testo al layout destro
        right_layout.addWidget(self.log_list)
        # Aggiungi il layout destro al layout orizzontale
        content_layout.addWidget(right_group, 1)  # Proporzione 1

        # Attiva il menu contestuale per il widget dei log
        self.log_list.setContextMenuPolicy(Qt.CustomContextMenu)
        self.log_list.customContextMenuRequested.connect(self.show_context_menu) 

        # Imposta l'allineamento in alto per il gruppo di sinistra
        content_layout.setAlignment(left_group, Qt.AlignTop)

        # Aggiungi un layot orizontale con tre bottoni
        button_widget = QWidget()
        button_widget.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Fixed)
        button_layout = QHBoxLayout(button_widget)
        button_layout.setAlignment(Qt.AlignLeft)  # Allinea il layout a sinistra
        ## aggiungo tre bottoni per 1)Da definire 2)Configura 3)Reset dati
        # Pulsante per la configurazione
        self.config_button = QPushButton("Configura")
        self.config_button.setFixedWidth(80)
        self.config_button.clicked.connect(self.on_config_clicked)
        button_layout.addWidget(self.config_button)
        # Pulsante per il reset dei dati
        self.reset_button = QPushButton("Reset")
        self.reset_button.setFixedWidth(80)
        self.reset_button.clicked.connect(self.on_reset_clicked)
        button_layout.addWidget(self.reset_button)
        # Pulsante per l'uscita
        self.exit_button = QPushButton("Esci")          
        self.exit_button.setFixedWidth(80)
        self.exit_button.clicked.connect(self.close)
        button_layout.addWidget(self.exit_button)

        # Aggiungi al layout principale
        main_layout.addLayout(content_layout)
        # Aggiungi il widget dei bottoni al layout del pannello sinistro
        main_layout.addWidget(button_widget)        
        
        # Status Bar
        self.statusBar = QStatusBar()
        self.setStatusBar(self.statusBar)

        # Aggiunge una progress bar alla status bar
        self.status_progress = QProgressBar()
        self.status_progress.setMaximumWidth(150)
        self.status_progress.setMaximumHeight(16)
        self.status_progress.setValue(0)
        self.statusBar.addPermanentWidget(self.status_progress)

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
        
        # Connessioni segnali
        self.start_date_picker.dateChanged.connect(self.on_date_changed)
        self.end_date_picker.dateChanged.connect(self.on_date_changed)

        # Carica la configurazione all'avvio
        self.config = self.load_config()
        # Assegna lo stato delle checkbox a variabili di classe
        self.estrai_adm_enabled = self.config["operations"].get("estrai_AdM", False)
        self.estrai_odm_enabled = self.config["operations"].get("estrai_OdM", False)
        self.elabora_xls_enabled = self.config["operations"].get("elabora_xls", False)

        # invio un msg alla finestra di log
        self.log_message("Sistema pronto!", 'info')

    # ------- definizione dei metodi per i bottoni --------

    def on_reset_clicked(self):
        """Resetta i campi di input e il log"""
        self.file_text.clear()
        self.start_date_picker.setDate(QDate.currentDate())
        self.end_date_picker.setDate(QDate.currentDate())
        self.log_list.clear()
        self.log_unified("Eseguito reset dell'applicativo", "info")
        self.start_button.setEnabled(False)
        self.excel_file_path = None  # Resetta il percorso del file Excel   

    def closeEvent(self, event):
        """Gestisce l'evento di chiusura della finestra"""
        self.log_unified("Applicazione terminata", "info")
        event.accept()
        # Chiude l'applicazione
        QApplication.quit()

    def on_config_clicked(self, message):
        pass

    def select_excel_file(self):
        # Apre il dialogo di selezione file
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Seleziona file Excel",
            "",
            "File Excel (*.xlsx *.xls *.xlsm)"
        )
        
        if file_path:
            # Estrai solo il nome del file dal percorso completo
            file_name = os.path.basename(file_path)
            self.log_unified(f"Selezionato file excel: {file_name}", "info")
            
            # Mostra il nome del file nel campo di testo
            self.file_text.setText(file_name)
            
            # Salva il percorso completo come attributo dell'oggetto
            self.excel_file_path = file_path
            
            # Verifica lo sheet
            self.log_unified("Verifico file excel", "loading", update_status=False, update_log=True)
            result, df = self.check_excel_file()
            if result:
                self.log_unified("Struttura file excel OK", "success", update_status=True, update_log=True)
                result, df_Norm = self.normalize_df(df)  # Salva il dataframe contenente la lista degli AdM per estrazioni IW29
                if not result:
                    # Se la normalizzazione fallisce, mostra un messaggio di errore
                    self.log_unified(f"Errore nella normalizzazione del file: {file_name}", "error")
                    return False                
                result, self.df_AdM = self.Estrai_AdM(df_Norm)  # Salva il dataframe contenente la lista degli AdM per estrazioni IW29
                if not result:
                    # Se l'estrazione fallisce, mostra un messaggio di errore
                    self.log_unified("Errore: Estrazione AdM fallita", "error")
                    return False
                msg = f"AdM estratti: {len(self.df_AdM)}"
                self.log_unified(msg, "info")
                # Verifica se ci sono OdM   
                result, self.df_OdM = self.Estrai_OdM(df_Norm)  # Salva il dataframe contenente la lista degli OdM per estrazioni IW39
                if not result:
                    # Se l'estrazione fallisce, mostra un messaggio di errore
                    self.log_unified("Errore: Estrazione OdM fallita", "error")
                    return False
                msg = f"OdM estratti: {len(self.df_OdM)}"
                self.log_unified(msg)
                # Abilita il pulsante di avvio
                self.start_button.setEnabled(True)
                self.log_unified("File excel caricato correttamente", "success")
                return True

            else:
                self.start_button.setEnabled(False)  # Disabilita il pulsante se lo sheet non esiste    
                msg = (f"Errore: File non valido - {file_name}")
                self.log_unified(msg, "error")
                # Elimina il nome del file nel campo di testo
                self.file_text.setText("")
                return False  
        else:
            self.log_unified("Nessun file selezionato", "warning")
            return False

    def load_config(self):
        """Carica la configurazione dal file o utilizza valori predefiniti"""
        default_config = constants.default_config        
        if not os.path.exists(self.config_file):
            msg = "File di configurazione non trovato, utilizzo valori predefiniti"
            self.log_unified("msg", 'warning')
            return default_config
        
        try:
            with open(self.config_file, 'r') as f:
                config = json.load(f)
            msg = "Configurazione caricata con successo"
            self.log_unified(msg, 'success')
            return config
        except Exception as e:
            # Se c'è un errore nel caricamento, restituisco la configurazione predefinita
            msg = f"Errore nel caricamento della configurazione: {str(e)}"
            self.log_unified(msg, 'warning')
            return default_config        
    
    def on_date_changed(self):
        start_date = self.start_date_picker.date()
        end_date = self.end_date_picker.date()
        
        if start_date > end_date:
            #self.statusBar.showMessage("Errore: La data di inizio è successiva alla data di fine!")
            self.start_button.setEnabled(False)
        else:
            #days = start_date.daysTo(end_date)
            #self.statusBar.showMessage(f"Intervallo selezionato: {days} giorni")
            self.start_button.setEnabled(True)
    
    def on_start_clicked(self):
        self.log_unified("Avvio estrazioni SAP")
        # Verifico se è stato selezionato un file
        self.log_unified("Verifico selezione file excel")
        if not hasattr(self, 'excel_file_path') or not self.excel_file_path:
            QMessageBox.warning(self, "Nessun file selezionato", 
                            "Seleziona un file Excel prima di procedere.")
            self.log_unified("Nessun file excel selezionato", "critical")
            return
        else:
            self.log_unified("Verifica file excel - OK", "success")
        # verifico la correttezza delle daate inserite
            # ricavo la data di inizio e fine
            start_date = self.start_date_picker.date()
            end_date = self.end_date_picker.date()
            # Verifica la validità dell'intervallo
            if not self.validate_date_range(start_date, end_date):
                self.log_unified("Errore nella verifica delle date inserite", "critical")
                return  # Esce dal metodo se la validazione fallisce

            # Ottieni la configurazione delle tecnologie
            tech_config = self.config.get("technologies", {})
            # Verifica la configurazione delle tecnologie
            if not self.validate_technology_config(tech_config):
                self.log_unified("Errore nella verifica dei codici tecnologia", "critical")
                return  # Esce dal metodo se la configurazione non è valida

            # estraggo i dati da SAP
            self.log_unified("Avvio estrazione SAP...")
            try:
                # Verifica della directory di salvataggio
                save_dir = self.config.get("save_directory", "")
                if not save_dir or not os.path.exists(save_dir):
                    self.log_unified("Errore: Directory di salvataggio non valida", "critical")
                    return
                with SAP_Connection.SAPGuiConnection() as sap:
                    if sap.is_connected():
                        session = sap.get_session()
                        if session:
                            self.log_unified("Connessione SAP attiva")
                            extractor = SAP_Transactions.SAPDataExtractor(session, self)
                            # Connetti il segnale a una lambda che gestisce i parametri, incluso origin
                            extractor.logMessage.connect(
                                lambda msg, lvl, upd_status, upd_log, min_time, origin, args, kwargs: 
                                    self.log_unified(msg, lvl, upd_status, upd_log, min_time, origin, *args, **kwargs)
                            )
                            # Estrazione dati AdM
                            self.log_unified("Estrazione dati IW29", "info", True, True, 0)
                            result, df_IW29 = extractor.extract_IW29(start_date, end_date, tech_config, self.df_AdM["idItem"] )
                            if not result:
                                self.log_unified("Errore: Estrazione IW29 fallita", "error", True, True, 0)
                                return
                            # Salva il DataFrame in un file Excel
                            output_file = os.path.join(save_dir, "IW29_AdM.xlsx")
                            if df_IW29.to_excel(output_file, index=False):
                                self.log_unified(f"File salvato in: {output_file}", "success", True, True, 0)
                            else:
                                self.log_unified("Errore: Salvataggio file IW29 fallito", "error", True, True, 0)
                                return
                            # Estrazione dati OdM
                            self.statusBar.showMessage("Estrazione dati IW39")
                            result, df_IW39 = extractor.extract_IW39(start_date, end_date, tech_config, self.df_OdM)
                            if not result:
                                self.log_unified("Errore: Estrazione IW39 fallita", "error", True, True, 0)
                                return
                            # Salva il DataFrame in un file Excel
                            output_file = os.path.join(save_dir, "IW39_OdM.xlsx")
                            if df_IW39.to_excel(output_file, index=False):
                                self.log_unified(f"File salvato in: {output_file}", "success", True, True, 0)
                            else:
                                self.log_unified("Errore: Salvataggio file IW39 fallito", "error", True, True, 0)
                                return
                            # Estrazione dati completata
                            self.log_unified("Estrazione completata con successo", "success", True, True, 0)
                    else:
                        self.log_unified("Connessione SAP NON attiva", "error", True, True, 0)
                        return
            except Exception as e:
                msg = (f"Estrazione dati SAP: Errore: {str(e)}")
                self.log_unified(msg, "error", True, True, 0)
                return           
            # ------------estrazione SAP completata---------------            
        
    def normalize_df(self, df) -> tuple[bool, pd.DataFrame | None]:
        """
        Estrae i dati AdM dal DataFrame fornito, elaborando i dati presenti nella colonna "idItem"
        
        Args:
            df (pd.DataFrame): DataFrame contenente i dati originali.
        
        Returns:
            tuple: Un tuple contenente un booleano e il DataFrame filtrato.
        """
        try:
            self.log_unified(" - Normalizzo la colonna idItem del DataFrame - ", "info", True, True, 0)
            self.log_unified(f"Presenti {len(df)} idItem nel DataFrame", "info", True, True, 0)
            # Crea un nuovo DataFrame con solo la colonna idItem
            id_items_df = df[["idItem"]].copy()
            # Verifica che il DataFrame non sia vuoto
            if id_items_df.empty:
                # Se il DataFrame è vuoto, mostra un messaggio di errore
                self.log_unified("Nessun elemento trovato nel file Excel", "error", True, True, 0)
                return False, None
                   
            # Rimuovi eventuali valori nulli
            id_items_df = id_items_df.dropna(subset=["idItem"])
            
            # Funzione per estrarre la parte prima del primo - o / e convertire in intero
            def extract_base_id(id_text):
                if not isinstance(id_text, str):
                    try:
                        # Se è già un numero, prova a convertirlo direttamente
                        return int(id_text)
                    except (ValueError, TypeError):
                        return None
                        
                # Cerca il primo trattino o slash
                dash_pos = id_text.find('-')
                slash_pos = id_text.find('/')
                
                # Determina quale carattere appare per primo (se presente)
                if dash_pos >= 0 and (slash_pos < 0 or dash_pos < slash_pos):
                    base_id = id_text[:dash_pos]
                elif slash_pos >= 0:
                    base_id = id_text[:slash_pos]
                else:
                    base_id = id_text  # Nessun trattino o slash trovato
                    
                # Converti in intero se possibile
                try:
                    return int(base_id)
                except ValueError:
                    # Se non può essere convertito, mantieni come stringa
                    self.log_unified(f"Impossibile convertire '{base_id}' in intero", "error", True, True, 0)
            
            # Applica la funzione a tutti i valori nella colonna idItem
            # Crea un nuovo DataFrame con entrambe le colonne per verificare il funzionamento
            new_df = id_items_df.copy()
            new_df['idItem_base'] = id_items_df['idItem'].apply(extract_base_id)
            id_items_df['idItem'] = id_items_df['idItem'].apply(extract_base_id)
            
            # Rimuovi eventuali duplicati
            id_items_df = id_items_df.drop_duplicates()
            
            # Resetta l'indice
            id_items_df = id_items_df.reset_index(drop=True)
            self.log_unified(f"Estratti {len(id_items_df)} idItem unici come valori interi", "success", True, True, 0)
            if (len(id_items_df) != len(df)):
                msg = f"{len(df) - len(id_items_df)} righe eliminate"
                self.log_unified(msg, "info", True, True, 0)
            return True, id_items_df
            
        except ValueError as ve:
            # Gestione specifica per ValueError (considerato un caso "atteso")
            self.log_unified(f"Normalizzazione fallita: {str(ve)}", "critical", True, True, 0)
            return False, None
            
        except Exception as e:
            self.log_unified(f"Errore nell'estrazione degli idItem: {str(e)}", "critical", True, True, 0)                        
            return False, None
    
    def Estrai_AdM(self, df) -> tuple[bool, pd.DataFrame | None]:
        """
        Filtra il DataFrame per restituire solo gli AdM ovvero le righe con idItem minori di 2000000000.
        
        Args:
            df (pandas.DataFrame): DataFrame contenente la colonna idItem con valori numerici
            
        Returns:
            pandas.DataFrame: Nuovo DataFrame contenente solo righe con idItem < 2000000000
        """
        try:
            msg = (f"Estrazione AdM - Presenti {len(df)} idItem nel DataFrame")
            self.log_unified(msg, "info", True, True, 0)
            # Assicurati che i valori siano numerici
            # Se i valori sono già numerici, questa operazione non avrà effetto
            # Se sono stringhe, tenta di convertirli in numeri
            try:
                df_numeric = df.copy()
                df_numeric["idItem"] = pd.to_numeric(df_numeric["idItem"], errors="coerce")
                # 'coerce' converte i valori non numerici in NaN
                
                # Rimuovi le righe con valori NaN (quelli che non era possibile convertire)
                df_numeric = df_numeric.dropna(subset=["idItem"])
                msg = (f"Pulizia df - Presenti {len(df_numeric)} idItem nel nuovo DataFrame")
                self.log_unified(msg, "info", True, True, 0)

                if (len(df_numeric) != len(df)):
                    msg = (f"Errore durante la conversione di: {len(df) - len(df_numeric)} righe")
                    self.log_unified(msg, "error", True, True, 0)

            except Exception as e:
                msg = (f"Errore nella conversione dei valori in numeri: {str(e)}")
                self.log_unified(msg, "error", True, True, 0)
                return False, None
                
            # Filtra il DataFrame per ottenere solo i valori minori di 2000000000
            AdM_df = df_numeric[df_numeric["idItem"] < 2000000000]
            
            # Resetta l'indice
            AdM_df = AdM_df.reset_index(drop=True)
            
            msg = (f"Filtrati {len(AdM_df)} record con idItem < 2000000000 su {len(df)} totali")
            self.log_unified(msg, "success", True, True, 0)
            return True, AdM_df
            
        except Exception as e:
            msg = (f"Errore nel filtro degli idItem: {str(e)}")
            self.log_unified(msg, "error", True, True, 0)
            return False, None

    def Estrai_OdM(self, df) -> tuple[bool, pd.DataFrame | None]:
        """
        Filtra il DataFrame per restituire solo gli OdM ovvero le righe con idItem maggiori di 2000000000.
        
        Args:
            df (pandas.DataFrame): DataFrame contenente la colonna idItem con valori numerici
            
        Returns:
            pandas.DataFrame: Nuovo DataFrame contenente solo righe con idItem > 2000000000
        """
        try:
            msg = (f"Estrazione OdM - Presenti {len(df)} idItem nel DataFrame")
            self.log_unified(msg, "info", True, True, 0)
            # Assicurati che i valori siano numerici
            # Se i valori sono già numerici, questa operazione non avrà effetto
            # Se sono stringhe, tenta di convertirli in numeri
            try:
                df_numeric = df.copy()
                df_numeric["idItem"] = pd.to_numeric(df_numeric["idItem"], errors="coerce")
                # 'coerce' converte i valori non numerici in NaN
                
                # Rimuovi le righe con valori NaN (quelli che non era possibile convertire)
                df_numeric = df_numeric.dropna(subset=["idItem"])
                msg = (f"Presenti {len(df_numeric)} idItem nel nuovo DataFrame")
                self.log_unified(msg, "info", True, True, 0)
                if (len(df_numeric) != len(df)):
                    msg = (f"Errore durante la conversione di: {len(df) - len(df_numeric)} righe")
                    self.log_unified(msg, "error", True, True, 0)

            except Exception as e:
                msg = (f"Errore nella conversione dei valori in numeri: {str(e)}")
                self.log_unified(msg, "error", True, True, 0)
                return False, None
                
            # Filtra il DataFrame per ottenere solo i valori maggiori di 2000000000
            OdM_df = df_numeric[df_numeric["idItem"] > 2000000000]
            
            # Resetta l'indice
            OdM_df = OdM_df.reset_index(drop=True)
            
            msg = (f"Filtrati {len(OdM_df)} record con idItem > 2000000000 su {len(df)} totali")
            self.log_unified(msg, "success", True, True, 0)
            return True, OdM_df
            
        except Exception as e:
            msg = (f"Errore nel filtro degli idItem: {str(e)}")
            self.log_unified(msg, "error", True, True, 0)
            return False, None

    def check_excel_file(self) -> tuple[bool, pd.DataFrame | None]:
        """
        Verifica se il file Excel selezionato contiene lo sheet richiesto
        e le colonne necessarie.
        
        Returns:
            bool: True se lo sheet e le colonne esistono, False altrimenti
        """
        try:
            # Verifica che il file sia stato selezionato
            if not hasattr(self, 'excel_file_path') or not self.excel_file_path:
                msg = "Errore: Nessun file selezionato"
                self.log_unified(msg, "error", True, True, 0)
                return False, None
            
            msg = (f"Verifica del file Excel: {self.excel_file_path}")
            self.log_unified(msg, "info", True, True, 0)

            # Ottieni la lista degli sheet nel file
            excel_file = pd.ExcelFile(self.excel_file_path)
            sheet_names = excel_file.sheet_names
            msg = (f"Sheet presenti nel file: {', '.join(sheet_names)}")
            self.log_unified(msg, "info", True, True, 0)
            
            # Definisci il nome dello sheet da cercare
            required_sheet = constants.required_sheet
            
            # Verifica se lo sheet esiste
            if required_sheet in sheet_names:
                msg = (f"Sheet '{required_sheet}' trovato nel file")
                self.log_unified(msg, "success", True, True, 0)
                # Leggi i dati dallo sheet
                df = pd.read_excel(self.excel_file_path, sheet_name=required_sheet)
                
                # Definisci le colonne richieste
                required_columns = constants.required_columns
                
                msg = (f"Colonne presenti nel file: {', '.join(df.columns)}")
                self.log_unified(msg, "info", True, True, 0)
                
                # Verifica se tutte le colonne richieste sono presenti
                missing_columns = [col for col in required_columns if col not in df.columns]
                
                if missing_columns:
                    # Alcune colonne sono mancanti
                    missing_cols_str = ", ".join(missing_columns)
                    msg = f"Errore: Colonne mancanti: {missing_cols_str}"
                    self.log_unified(msg, "error", True, True, 0)
                    msg = (f"Colonne disponibili: {', '.join(df.columns)}")
                    self.log_unified(msg, "error", True, True, 0)
                    return False, None
                else:
                    # Tutte le colonne sono presenti
                    rows_count = len(df)
                    msg = f"File valido: {rows_count} righe trovate"
                    self.log_unified(msg, "success", True, True, 0)
                    return True, df
            else:
                # Lo sheet non è presente
                msg = f"Errore: Lo sheet '{required_sheet}' non è presente nel file"
                self.log_unified(msg, "error", True, True, 0)
                return False, None
                
        except Exception as e:
            msg = f"Errore nella lettura del file Excel: {str(e)}"
            self.log_unified(msg, "error", True, True, 0)
            return False, None
    
    def on_config_clicked(self):
        """Apre la finestra di configurazione"""
        self.statusBar.showMessage("Apertura configurazione...")
        
        config_dialog = ConfigDialog(
                                        config_file = self.config_file, 
                                        parent = self
                                    )
        # Connetti il segnale di log attività al gestore nella finestra principale
        config_dialog.activityLogged.connect(self.log_message)

        if config_dialog.exec_():
            # Ricarica la configurazione se è stata salvata
            self.config = self.load_config()
            # Assegna lo stato delle checkbox a variabili di classe
            self.estrai_adm_enabled = self.config["operations"].get("estrai_AdM", False)
            self.estrai_odm_enabled = self.config["operations"].get("estrai_OdM", False)
            self.elabora_xls_enabled = self.config["operations"].get("elabora_xls", False)            
            self.log_unified("Configurazione aggiornata", "info")
        else:
            self.log_unified("Configurazione non modificata", "info")

    def validate_date_range(self, start_date, end_date):
        """
        Verifica che l'intervallo delle date sia valido e mostra un messaggio di errore se necessario.
        
        Args:
            start_date: QDate di inizio
            end_date: QDate di fine
            
        Returns:
            bool: True se le date sono valide, False altrimenti
        """
        msg=(f"Verifica date inserite: {start_date.toString('dd/MM/yyyy')} - {end_date.toString('dd/MM/yyyy')}")
        self.log_unified(msg)
        # Verifica che la data di inizio sia precedente o uguale alla data di fine
        if start_date > end_date:
            error_message = "La data di inizio non può essere successiva alla data di fine"
            self.log_unified(error_message, "error")            
            QMessageBox.warning(self, "Errore nell'intervallo date", error_message)
            return False
        
        # Verifica che la data di inizio non sia uguale alla data di fine
        if start_date == end_date:
            error_message = "La data di inizio non può essere uguale alla data di fine"
            self.log_unified(error_message, "error")            
            QMessageBox.warning(self, "Errore nell'intervallo date", error_message)
            return False
        
        # Calcola la differenza in giorni
        days_difference = start_date.daysTo(end_date)
        
        # Verifica che l'intervallo non superi un anno (365 o 366 giorni)
        max_days = 366  # Per considerare anche gli anni bisestili
        if days_difference > max_days:
            error_message = f"L'intervallo di date non può superare un anno ({max_days} giorni)"
            self.log_unified(error_message, "error")
            QMessageBox.warning(self, "Intervallo troppo ampio", error_message)
            return False
        
        # Tutte le verifiche sono state superate
        self.log_unified("Verifica date inserite - OK", "success")
        return True     

    def validate_technology_config(self, tech_config):
        """
        Verifica che le tecnologie siano configurate correttamente.
        
        Returns:
            bool: True se la configurazione è valida, False altrimenti
        """
        self.log_unified(f"Verifica tecnologie configurate: {tech_config}")

        # Verifica che tech_config non sia vuoto
        if not tech_config:
            error_message = "Nessuna tecnologia configurata. Utilizzare la finestra di configurazione."
            QMessageBox.warning(self, "Configurazione mancante", error_message)
            self.log_unified("Nessuna tecnologia configurata", "error")
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
            self.log_unified(error_message, "error")
            return False
        
        # Se arriviamo qui, la configurazione è valida
        self.log_unified("Verifica tecnologie configurate - OK", "success")
        return True
    # --------------------------------------------------------------------------------------------------------
    # Funzioni per scrivere messaggi nel log nel QListWidget con icona appropriata
    # --------------------------------------------------------------------------------------------------------
    def log_message(self, message, icon_type='info'):
        """
        Aggiunge un messaggio al log con un'icona Qt
        """
        item = QListWidgetItem(message)
        
        # Imposta l'icona in base al tipo
        if icon_type == 'info':
            item.setIcon(self.style().standardIcon(QStyle.SP_MessageBoxInformation))
        elif icon_type == 'error':
            item.setIcon(self.style().standardIcon(QStyle.SP_MessageBoxCritical))
        elif icon_type == 'success':
            item.setIcon(self.style().standardIcon(QStyle.SP_DialogApplyButton))
        elif icon_type == 'warning':
            item.setIcon(self.style().standardIcon(QStyle.SP_MessageBoxWarning))
        elif icon_type == 'loading':
            item.setIcon(self.style().standardIcon(QStyle.SP_BrowserReload))
        
        self.log_list.addItem(item)
        self.log_list.scrollToBottom()

    # --------------------------------------------------------------------------------------------------------
    # Metodo unificato per gestire log attraverso tutti i canali disponibili.
    # --------------------------------------------------------------------------------------------------------

    def log_unified(self, message, level="info", update_status=True, update_log=True, min_display_seconds=None, origin="main", *args, **kwargs):
        """
        Metodo unificato per gestire log attraverso tutti i canali disponibili.
        
        Args:
            message: Il messaggio da registrare
            level: Livello del log ('info', 'warning', 'error', ecc.)
            update_status: Se aggiornare la statusBar
            update_log: Se aggiornare il widget di log visuale
            min_display_seconds: Tempo minimo di visualizzazione
            origin: Origine del messaggio per evitare duplicazioni
            *args, **kwargs: Argomenti per la formattazione del messaggio
        """
        # Formatta il messaggio se necessario
        try:
            if kwargs and not args:
                formatted_message = message.format(**kwargs)
            elif args:
                formatted_message = message % args
            else:
                formatted_message = message
        except Exception as e:
            formatted_message = f"{message} (Errore formattazione: {str(e)})"
        
        # 1. Mappa il livello ai tipi di icona disponibili per log_message
        icon_map = {
            "debug": "info",      # debug usa l'icona info
            "info": "info",       # icona standard informativa
            "success": "success", # icona segno di spunta/approvazione
            "warning": "warning", # icona triangolo di avviso
            "error": "error",     # icona X rossa critica
            "critical": "error",  # anche critical usa l'icona error
            "loading": "loading"  # icona di caricamento/aggiornamento
        }
        icon_type = icon_map.get(level, "info")  # default a info se il livello non è mappato
        
        # 2. Registra nel QListWidget di log con l'icona appropriata (se richiesto)
        if update_log:
            self.log_message(formatted_message, icon_type)
        
        # 3. Registra nel logger di sistema con il livello appropriato, ma SOLO se l'origine è "main"
        # Questo evita la duplicazione dei log
        if origin == "main":
            # Mappa i livelli speciali alle funzioni del logger
            logger_level_map = {
                "success": "info",    # success non è un livello standard, usa info
                "loading": "info"     # loading non è un livello standard, usa info
            }
            # Ottieni il livello standard corrispondente o usa lo stesso se è già standard
            logger_level = logger_level_map.get(level, level)
            logger_method = getattr(logger, logger_level) if hasattr(logger, logger_level) else logger.info
            logger_method(formatted_message)
        
        # 4. Gestione della statusBar basata sul livello e tempo minimo di visualizzazione
        if update_status:
            # Inizializza gli attributi di stato se necessario
            if not hasattr(self, '_last_status_level'):
                self._last_status_level = ""
            if not hasattr(self, '_last_status_time'):
                self._last_status_time = 0
            if not hasattr(self, '_last_status_min_time'):
                self._last_status_min_time = 0
                
            # Livelli in ordine crescente di importanza
            level_importance = ["debug", "info", "success", "loading", "warning", "error", "critical"]
            
            # Tempi minimi di visualizzazione predefiniti per livello (in secondi)
            default_min_times = {
                "debug": 2,
                "info": 3,
                "success": 5,
                "loading": 2,
                "warning": 8,
                "error": 10,
                "critical": 15
            }
            
            # Usa il tempo minimo specificato o il valore predefinito per il livello
            if min_display_seconds is None:
                min_display_seconds = default_min_times.get(level, 3)
            
            current_time = time.time()
            
            # Calcola gli indici di importanza (posizione nell'array level_importance)
            try:
                current_level_idx = level_importance.index(level)
            except ValueError:
                current_level_idx = 0  # Default a bassa importanza se livello sconosciuto
                
            try:
                last_level_idx = level_importance.index(self._last_status_level)
            except ValueError:
                last_level_idx = 0
                
            # Calcola quanto tempo è passato dall'ultimo aggiornamento
            elapsed_time = current_time - self._last_status_time
            
            # Determina se aggiornare la statusBar:
            # 1. Se è passato il tempo minimo dell'ultimo messaggio, OPPURE
            # 2. Se il nuovo livello è più importante dell'ultimo, OPPURE
            # 3. Se il livello è lo stesso dell'ultimo (aggiornamento incrementale)
            update_condition = (
                elapsed_time >= self._last_status_min_time or
                current_level_idx > last_level_idx or
                level == self._last_status_level
            )
            
            if update_condition:
                # Per i vari livelli, prependi il tipo di messaggio se necessario
                prefix_map = {
                    "error": "Errore: ",
                    "critical": "Errore critico: ",
                    "warning": "Attenzione: ",
                    "loading": "Caricamento: "
                }
                prefix = prefix_map.get(level, "")
                status_message = f"{prefix}{formatted_message}"
                
                self.statusBar.showMessage(status_message)
                
                # Aggiorna le informazioni sull'ultimo stato
                self._last_status_level = level
                self._last_status_time = current_time
                self._last_status_min_time = min_display_seconds        

    # --------------------------------------------------------------------------------------------------------
    # Funzioni per mostrare un menu contestuale x copiare i dati
    # --------------------------------------------------------------------------------------------------------
    def show_context_menu(self, position):
        # Crea menu contestuale
        context_menu = QMenu()
        
        # Aggiungi l'azione "Copia"
        copy_action = QAction("Copia elemento", self)
        copy_action.triggered.connect(self.copy_selected_items)
        context_menu.addAction(copy_action)
        
        # Aggiungi l'azione "Copia tutto"
        copy_all_action = QAction("Copia tutto", self)
        copy_all_action.triggered.connect(self.copy_all_items)
        context_menu.addAction(copy_all_action)
        
        # Mostra il menu contestuale alla posizione corrente del cursore
        context_menu.exec_(QCursor.pos())

    def copy_selected_items(self):
        # Copia solo gli elementi selezionati
        selected_items = self.log_list.selectedItems()
        if selected_items:
            text = "\n".join(item.text() for item in selected_items)
            QApplication.clipboard().setText(text)
            print("Elementi selezionati copiati negli appunti")        

    def copy_all_items(self):
        # Copia tutti gli elementi
        all_items = []
        for i in range(self.log_list.count()):
            all_items.append(self.log_list.item(i).text())
        
        text = "\n".join(all_items)
        QApplication.clipboard().setText(text)
        print("Tutti gli elementi copiati negli appunti")  

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())