import sys
import os
import json
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                            QHBoxLayout, QLabel, QDateEdit, QTextEdit, 
                            QPushButton, QStatusBar, QMessageBox)
from PyQt5.QtCore import QDate, Qt
import SAP_Connection
import SAP_Transactions

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

        # Imposta i flag della finestra per mostrare solo il pulsante di chiusura
        self.setWindowFlags(self.windowFlags() & ~Qt.WindowMaximizeButtonHint & ~Qt.WindowMinimizeButtonHint)
            
        self.setWindowTitle("Estrai Dati KPI OFA")
        self.setWindowIconText("KPI OFA")
        self.setGeometry(100, 100, 160, 400)  # x, y, width, height
        
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
        self.text_edit = QTextEdit()
        self.text_edit.setPlaceholderText("Incolla qui la lista AdM/OdM...")
        main_layout.addWidget(self.text_edit)
        
        # Riga 4 - Bottoni
        row4_layout = QHBoxLayout()
        self.start_button = QPushButton("Avvia")
        self.start_button.clicked.connect(self.on_start_clicked)
        self.start_button.setFixedWidth(80)
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

    def load_config(self):
        """Carica la configurazione dal file o utilizza valori predefiniti"""
        default_config = {
            "save_directory": os.path.expanduser("~"),
            "technologies": {
                "BESS": ["ITE", "USE", "CLE"],
                "SOLAR": ["ITS", "USS", "CLS", "BRS", "COS", "MXS", "PAS", "ZAS", "ESS", "ZMS"],
                "WIND": ["ITW", "USW", "CLW", "BRW", "CAW", "MXW", "ZAW", "ESW"]
            }
        }
        
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
        text_content = self.text_edit.toPlainText()
        # verifico se la finestra di testo è vuota
        if not text_content:
            self.statusBar.showMessage("Errore: Nessun dato inserito nella finestra di testo")
            return
        else:
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
                            df_IW29 = extractor.extract_IW29(start_date, end_date, tech_config)
                            df_IW39 = extractor.extract_IW39(start_date, end_date, tech_config)
                            
                            self.statusBar.showMessage("Estrazione completata con successo")
                    else:
                        self.statusBar.showMessage("Connessione SAP NON attiva")
                        return
            except Exception as e:
                self.statusBar.showMessage(f"Estrazione dati SAP: Errore: {str(e)}")
                return           
            # ------------estrazione SAP completata---------------            
        
        lines = text_content.strip().split('\n')
        self.statusBar.showMessage(f"Elaborazione avviata con {len(lines)} valori")
        # Qui puoi aggiungere la logica per elaborare i dati
    
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
            rror_message = f"L'intervallo di date non può superare un anno ({max_days} giorni)"
            logger.error(error_message)
            QMessageBox.warning(self, "Intervallo troppo ampio", rror_message)
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