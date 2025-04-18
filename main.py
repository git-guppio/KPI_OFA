import sys
import os
import json
import pandas as pd
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                            QHBoxLayout, QLabel, QDateEdit, QFileDialog, QLineEdit,
                            QPushButton, QStatusBar, QMessageBox)
from PyQt5.QtCore import QDate, Qt
import SAP_Connection
import SAP_Transactions
import Config.constants as constants
import openpyxl
import logging

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
        self.config_file = "config.json"
        # Carica la configurazione all'avvio
        self.config = self.load_config()
        # Dataframe contenenti i dati dal file Excel attraverso la finestra di selezione file
        self.df_AdM = None
        self.df_OdM = None
        self.excel_file_path = None  # Percorso del file Excel selezionato

        # Imposta i flag della finestra per mostrare solo il pulsante di chiusura
        self.setWindowFlags(self.windowFlags() & ~Qt.WindowMaximizeButtonHint & ~Qt.WindowMinimizeButtonHint)
            
        self.setWindowTitle("Estrai Dati KPI OFA")
        self.setWindowIconText("KPI OFA")
        self.setGeometry(100, 100, 180, 168)  # x, y, width, height
        
        # Widget centrale
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        # Layout principale
        main_layout = QVBoxLayout(central_widget)
        
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
        main_layout.addLayout(row1_layout)
        
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
        main_layout.addLayout(row2_layout)
        
        # Riga 3 - Finestra di testo
        """     
        self.text_edit = QTextEdit()
        self.text_edit.setPlaceholderText("Incolla qui la lista AdM/OdM...")
        main_layout.addWidget(self.text_edit)
        """

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

        # Aggiungi il layout al layout principale
        main_layout.addLayout(file_layout)
        
        # Riga 4 - Bottoni
        row4_layout = QHBoxLayout()
        self.start_button = QPushButton("Avvia")
        self.start_button.clicked.connect(self.on_start_clicked)
        self.start_button.setFixedWidth(80)
        self.start_button.setEnabled(False)  # Disabilita il pulsante se lo sheet non esiste
        self.config_button = QPushButton("Configura")
        self.config_button.clicked.connect(self.on_config_clicked)
        self.config_button.setFixedWidth(80)
        row4_layout.addWidget(self.start_button)
        row4_layout.addWidget(self.config_button)
        row4_layout.addStretch()
        main_layout.addLayout(row4_layout)
        
        # Status Bar
        self.statusBar = QStatusBar()
        self.setStatusBar(self.statusBar)
        self.statusBar.showMessage("Pronto")
        
        # Connessioni segnali
        self.start_date_picker.dateChanged.connect(self.on_date_changed)
        self.end_date_picker.dateChanged.connect(self.on_date_changed)

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
            
            # Mostra il nome del file nel campo di testo
            self.file_text.setText(file_name)
            
            # Salva il percorso completo come attributo dell'oggetto
            self.excel_file_path = file_path
            
            # Verifica lo sheet
            result, df = self.check_excel_file()
            if result:
                result, df_Norm = self.normalize_df(df)  # Salva il dataframe contenente la lista degli AdM per estrazioni IW29
                if not result:
                    self.statusBar.showMessage("Errore: Normalizzazione df fallita")
                    return False                
                result, self.df_AdM = self.Estrai_AdM(df_Norm)  # Salva il dataframe contenente la lista degli AdM per estrazioni IW29
                if not result:
                    self.statusBar.showMessage("Errore: Estrazione AdM fallita")
                    return False
                msg = f"AdM estratti: {len(self.df_AdM)}"
                logger.info(msg)
                self.statusBar.showMessage(msg)
                # Verifica se ci sono OdM   
                result, self.df_OdM = self.Estrai_OdM(df_Norm)  # Salva il dataframe contenente la lista degli OdM per estrazioni IW39
                if not result:
                    self.statusBar.showMessage("Errore: Estrazione OdM fallita")
                    return False
                msg = f"OdM estratti: {len(self.df_AdM)}"
                logger.info(msg)
                self.statusBar.showMessage(msg)

            else:
                self.start_button.setEnabled(False)  # Disabilita il pulsante se lo sheet non esiste    
                self.statusBar.showMessage(f"Errore: File non valido - {file_name}")

                
                
        else:
            self.statusBar.showMessage("Nessun file selezionato")

    def load_config(self):
        """Carica la configurazione dal file o utilizza valori predefiniti"""
        default_config = constants.default_config        
        if not os.path.exists(self.config_file):
            logger.info("File di configurazione non trovato, utilizzo valori predefiniti")
            return default_config
        
        try:
            with open(self.config_file, 'r') as f:
                config = json.load(f)
            logger.info("Configurazione caricata con successo")
            return config
        except Exception as e:
            logger.error(f"Errore nel caricamento della configurazione: {str(e)}")
            return default_config        
    
    def on_date_changed(self):
        start_date = self.start_date_picker.date()
        end_date = self.end_date_picker.date()
        
        if start_date > end_date:
            self.statusBar.showMessage("Errore: La data di inizio è successiva alla data di fine!")
            self.start_button.setEnabled(False)
        else:
            days = start_date.daysTo(end_date)
            self.statusBar.showMessage(f"Intervallo selezionato: {days} giorni")
            self.start_button.setEnabled(True)
    
    def on_start_clicked(self):
        # Verifico se è stato selezionato un file
        if not hasattr(self, 'excel_file_path') or not self.excel_file_path:
            QMessageBox.warning(self, "Nessun file selezionato", 
                            "Seleziona un file Excel prima di procedere.")
            self.statusBar.showMessage("Errore: Nessun file selezionato")
            return
        else:
        # Verifico la struttura del file Excel e carico i dati in un DataFrame
        #     
        # verifico la correttezza delle daate inserite
            # ricavo la data di inizio e fine
            start_date = self.start_date_picker.date()
            end_date = self.end_date_picker.date()
            # Verifica la validità dell'intervallo
            if not self.validate_date_range(start_date, end_date):
                return  # Esce dal metodo se la validazione fallisce

            # Ottieni la configurazione delle tecnologie
            tech_config = self.config.get("technologies", {})
            # Verifica la configurazione delle tecnologie
            if not self.validate_technology_config(tech_config):
                return  # Esce dal metodo se la configurazione non è valida

            # estraggo i dati da SAP
            self.statusBar.showMessage("Avvio estrazione...")
            try:
                # Verifica della directory di salvataggio
                save_dir = self.config.get("save_directory", "")
                if not save_dir or not os.path.exists(save_dir):
                    self.statusBar.showMessage("Errore: Directory di salvataggio non valida")
                    return
                with SAP_Connection.SAPGuiConnection() as sap:
                    if sap.is_connected():
                        session = sap.get_session()
                        if session:
                            self.statusBar.showMessage("Connessione SAP attiva")
                            extractor = SAP_Transactions.SAPDataExtractor(session)
                            self.statusBar.showMessage("Estrazione dati IW29")
                            result, df_IW29 = extractor.extract_IW29(start_date, end_date, tech_config)
                            self.statusBar.showMessage("Estrazione dati IW39")
                            df_IW39 = extractor.extract_IW39(start_date, end_date, tech_config)
                            
                            self.statusBar.showMessage("Estrazione completata con successo")
                    else:
                        self.statusBar.showMessage("Connessione SAP NON attiva")
                        return
            except Exception as e:
                self.statusBar.showMessage(f"Estrazione dati SAP: Errore: {str(e)}")
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
            logger.info(" - Normalizzo la colonna idItem del DataFrame - ")
            logger.info(f"Presenti {len(df)} idItem nel DataFrame")
            # Crea un nuovo DataFrame con solo la colonna idItem
            id_items_df = df[["idItem"]].copy()
            # Verifica che il DataFrame non sia vuoto
            if id_items_df.empty:
                logger.warning("Nessun elemento trovato nel file Excel")
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
                    logger.warning(f"Impossibile convertire '{base_id}' in intero")
            
            # Applica la funzione a tutti i valori nella colonna idItem
            id_items_df['idItem'] = id_items_df['idItem'].apply(extract_base_id)
            
            # Rimuovi eventuali duplicati
            id_items_df = id_items_df.drop_duplicates()
            
            # Resetta l'indice
            id_items_df = id_items_df.reset_index(drop=True)
            
            logger.info(f"Estratti {len(id_items_df)} idItem unici come valori interi")
            if (len(id_items_df) != len(df)):
                msg = f"{len(df) - len(id_items_df)} righe non convertite"
                logger.warning(f"Normalizzazione: {msg}")
            return True, id_items_df
            
        except ValueError as ve:
            # Gestione specifica per ValueError (considerato un caso "atteso")
            logger.warning(f"Normalizzazione fallita: {str(ve)}")
            return False, None
            
        except Exception as e:
            logger.exception(f"Errore nell'estrazione degli idItem: {str(e)}")
            return False, None
    
    def Estrai_AdM(self, df) -> tuple[bool, pd.DataFrame | None]:
        """
        Filtra il DataFrame per restituire solo gli AdM ovvero le righe con idItem maggiori di 2000000000.
        
        Args:
            df (pandas.DataFrame): DataFrame contenente la colonna idItem con valori numerici
            
        Returns:
            pandas.DataFrame: Nuovo DataFrame contenente solo righe con idItem > 2000000000
        """
        try:
            logger.info(" - Estrazione AdM - ")
            logger.info(f"Presenti {len(df)} idItem nel DataFrame")
            # Assicurati che i valori siano numerici
            # Se i valori sono già numerici, questa operazione non avrà effetto
            # Se sono stringhe, tenta di convertirli in numeri
            try:
                df_numeric = df.copy()
                df_numeric["idItem"] = pd.to_numeric(df_numeric["idItem"], errors="coerce")
                # 'coerce' converte i valori non numerici in NaN
                
                # Rimuovi le righe con valori NaN (quelli che non era possibile convertire)
                df_numeric = df_numeric.dropna(subset=["idItem"])
                logger.info(f"Presenti {len(df_numeric)} idItem nel nuovo DataFrame")

                if (len(df_numeric) != len(df)):
                    logger.warning(f"Errore durante la conversione di: {len(df) - len(df_numeric)} righe")

            except Exception as e:
                logger.error(f"Errore nella conversione dei valori in numeri: {str(e)}")
                return False, None
                
            # Filtra il DataFrame per ottenere solo i valori maggiori di 2000000000
            AdM_df = df_numeric[df_numeric["idItem"] > 2000000000]
            
            # Resetta l'indice
            AdM_df = AdM_df.reset_index(drop=True)
            
            logger.info(f"Filtrati {len(AdM_df)} record con idItem > 2000000000 su {len(df)} totali")
            
            return True, AdM_df
            
        except Exception as e:
            logger.exception(f"Errore nel filtro degli idItem: {str(e)}")
            return False, None

    def Estrai_OdM(self, df) -> tuple[bool, pd.DataFrame | None]:
        """
        Filtra il DataFrame per restituire solo gli OdM ovvero le righe con idItem minori di 2000000000.
        
        Args:
            df (pandas.DataFrame): DataFrame contenente la colonna idItem con valori numerici
            
        Returns:
            pandas.DataFrame: Nuovo DataFrame contenente solo righe con idItem < 2000000000
        """
        try:
            logger.info(" - Estrazione OdM - ")
            logger.info(f"Presenti {len(df)} idItem nel DataFrame")
            # Assicurati che i valori siano numerici
            # Se i valori sono già numerici, questa operazione non avrà effetto
            # Se sono stringhe, tenta di convertirli in numeri
            try:
                df_numeric = df.copy()
                df_numeric["idItem"] = pd.to_numeric(df_numeric["idItem"], errors="coerce")
                # 'coerce' converte i valori non numerici in NaN
                
                # Rimuovi le righe con valori NaN (quelli che non era possibile convertire)
                df_numeric = df_numeric.dropna(subset=["idItem"])
                logger.info(f"Presenti {len(df_numeric)} idItem nel nuovo DataFrame")

                if (len(df_numeric) != len(df)):
                    logger.warning(f"Errore durante la conversione di: {len(df) - len(df_numeric)} righe")

            except Exception as e:
                logger.error(f"Errore nella conversione dei valori in numeri: {str(e)}")
                return False, None
                
            # Filtra il DataFrame per ottenere solo i valori maggiori di 2000000000
            OdM_df = df_numeric[df_numeric["idItem"] < 2000000000]
            
            # Resetta l'indice
            OdM_df = OdM_df.reset_index(drop=True)
            
            logger.info(f"Filtrati {len(OdM_df)} record con idItem < 2000000000 su {len(df)} totali")
            
            return True, OdM_df
            
        except Exception as e:
            logger.exception(f"Errore nel filtro degli idItem: {str(e)}")
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
                error_msg = "Errore: Nessun file selezionato"
                logger.error(error_msg)
                self.statusBar.showMessage(error_msg)
                return False, None
            
            logger.info(f"Verifica del file Excel: {self.excel_file_path}")
            
            # Ottieni la lista degli sheet nel file
            excel_file = pd.ExcelFile(self.excel_file_path)
            sheet_names = excel_file.sheet_names
            logger.debug(f"Sheet presenti nel file: {', '.join(sheet_names)}")
            
            # Definisci il nome dello sheet da cercare
            required_sheet = constants.required_sheet
            
            # Verifica se lo sheet esiste
            if required_sheet in sheet_names:
                logger.info(f"Sheet '{required_sheet}' trovato nel file")
                
                # Leggi i dati dallo sheet
                df = pd.read_excel(self.excel_file_path, sheet_name=required_sheet)
                
                # Definisci le colonne richieste
                required_columns = constants.required_columns
                
                logger.debug(f"Colonne presenti nel file: {', '.join(df.columns)}")
                
                # Verifica se tutte le colonne richieste sono presenti
                missing_columns = [col for col in required_columns if col not in df.columns]
                
                if missing_columns:
                    # Alcune colonne sono mancanti
                    missing_cols_str = ", ".join(missing_columns)
                    error_msg = f"Errore: Colonne mancanti: {missing_cols_str}"
                    logger.error(error_msg)
                    self.statusBar.showMessage(error_msg)
                    logger.debug(f"Colonne disponibili: {', '.join(df.columns)}")
                    return False, None
                else:
                    # Tutte le colonne sono presenti
                    rows_count = len(df)
                    success_msg = f"File valido: {rows_count} righe trovate"
                    logger.info(success_msg)
                    self.statusBar.showMessage(success_msg)
                    return True, df
            else:
                # Lo sheet non è presente
                error_msg = f"Errore: Lo sheet '{required_sheet}' non è presente nel file"
                logger.error(error_msg)
                self.statusBar.showMessage(error_msg)
                logger.debug(f"Sheet disponibili: {', '.join(sheet_names)}")
                return False, None
                
        except Exception as e:
            error_msg = f"Errore nella lettura del file Excel: {str(e)}"
            logger.exception(error_msg)  # logger.exception include anche il traceback completo
            self.statusBar.showMessage(error_msg)
            return False, None
    
    def on_config_clicked(self):
        """Apre la finestra di configurazione"""
        self.statusBar.showMessage("Apertura configurazione...")
        
        config_dialog = ConfigDialog(self)
        if config_dialog.exec_():
            # Ricarica la configurazione se è stata salvata
            self.config = self.load_config()
            self.statusBar.showMessage("Configurazione aggiornata")
        else:
            self.statusBar.showMessage("Configurazione annullata")

    def validate_date_range(self, start_date, end_date):
        """
        Verifica che l'intervallo delle date sia valido e mostra un messaggio di errore se necessario.
        
        Args:
            start_date: QDate di inizio
            end_date: QDate di fine
            
        Returns:
            bool: True se le date sono valide, False altrimenti
        """
        logger.info(f"Verifica date inserite: {start_date} - {end_date}")
        # Verifica che la data di inizio sia precedente o uguale alla data di fine
        if start_date > end_date:
            error_message = "La data di inizio non può essere successiva alla data di fine"
            logger.error(error_message)
            QMessageBox.warning(self, "Errore nell'intervallo date", error_message)
            return False
        
        # Calcola la differenza in giorni
        days_difference = start_date.daysTo(end_date)
        
        # Verifica che l'intervallo non superi un anno (365 o 366 giorni)
        max_days = 366  # Per considerare anche gli anni bisestili
        if days_difference > max_days:
            error_message = f"L'intervallo di date non può superare un anno ({max_days} giorni)"
            logger.error(error_message)
            QMessageBox.warning(self, "Intervallo troppo ampio", error_message)
            return False
        
        # Tutte le verifiche sono state superate
        logger.info(f"Verifica date - OK")
        return True     

    def validate_technology_config(self, tech_config):
        """
        Verifica che le tecnologie siano configurate correttamente.
        
        Returns:
            bool: True se la configurazione è valida, False altrimenti
        """
        logger.info(f"Verifica tecnologie configurate: {tech_config}")

        # Verifica che tech_config non sia vuoto
        if not tech_config:
            error_message = "Nessuna tecnologia configurata. Utilizzare la finestra di configurazione."
            QMessageBox.warning(self, "Configurazione mancante", error_message)
            logger.error(error_message)
            self.statusBar.showMessage(error_message)
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
            logger.error(error_message)
            self.statusBar.showMessage(error_message)
            return False
        
        # Se arriviamo qui, la configurazione è valida
        logger.info(f"Verifica tecnologie configurate - OK")
        return True

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())