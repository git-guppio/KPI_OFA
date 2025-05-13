import os
import json
from PyQt5.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QLabel, 
                            QLineEdit, QPushButton, QFileDialog, QTabWidget,
                            QListWidget, QGroupBox, QMessageBox, QWidget,
                            QCheckBox)
from PyQt5.QtCore import Qt, pyqtSignal

class ConfigDialog(QDialog):

    # Definisci un segnale personalizzato per il log
    activityLogged = pyqtSignal(str, str)  # (message, icon_type)

    def __init__(self, config_file="config.json", parent=None):
        super().__init__(parent)
        self.setWindowTitle("Configurazione")
        self.setMinimumSize(500, 400)
        self.config_file = config_file
        
        # Layout principale
        main_layout = QVBoxLayout(self)
        
        # Sezione directory
        dir_group = QGroupBox("Directory di salvataggio")
        dir_layout = QHBoxLayout()
        
        self.dir_input = QLineEdit()
        self.dir_input.setPlaceholderText("Seleziona directory per salvare i dati...")
        self.dir_input.setMinimumWidth(300)
        
        browse_button = QPushButton("Sfoglia...")
        browse_button.clicked.connect(self.browse_directory)
        
        dir_layout.addWidget(self.dir_input)
        dir_layout.addWidget(browse_button)
        dir_group.setLayout(dir_layout)
        main_layout.addWidget(dir_group)

        # configura operazioni da eseguire
        op_group = QGroupBox("Operazioni da eseguire")
        op_layout = QHBoxLayout()
        # Crea le checkbox
        self.cb_estrai_adm = QCheckBox("Estrai AdM")
        self.cb_estrai_odm = QCheckBox("Estrai OdM")
        self.cb_elabora_xls = QCheckBox("Elabora File XLS")

        # Aggiungi le checkbox al layout
        op_layout.addWidget(self.cb_estrai_adm)
        op_layout.addWidget(self.cb_estrai_odm)
        op_layout.addWidget(self.cb_elabora_xls)
        # Imposta il layout del gruppo
        op_group.setLayout(op_layout)
        main_layout.addWidget(op_group)
        
        # Sezione tecnologie con tab
        tech_group = QGroupBox("Configurazione Tecnologie")
        tech_layout = QVBoxLayout()
        
        tab_widget = QTabWidget()
        
        # Definizione delle tecnologie e dei relativi prefissi
        self.technologies = {
            "BESS": ["ITE", "USE", "CLE"],
            "SOLAR": ["ITS", "USS", "CLS", "BRS", "COS", "MXS", "PAS", "ZAS", "ESS", "ZMS"],
            "WIND": ["ITW", "USW", "CLW", "BRW", "CAW", "MXW", "ZAW", "ESW"]
        }
        
        # Creazione di una tab per ciascuna tecnologia
        self.list_widgets = {}
        for tech, prefixes in self.technologies.items():
            tab = QWidget()
            tab_layout = QVBoxLayout(tab)
            
            # Lista dei prefissi
            self.list_widgets[tech] = QListWidget()
            
            # Aggiungi i prefissi alla lista
            self.list_widgets[tech].addItems(prefixes)
            
            # Pulsanti per aggiungere/rimuovere prefissi
            buttons_layout = QHBoxLayout()
            add_button = QPushButton("Aggiungi")
            add_button.clicked.connect(lambda checked, t=tech: self.add_prefix(t))
            
            remove_button = QPushButton("Rimuovi")
            remove_button.clicked.connect(lambda checked, t=tech: self.remove_prefix(t))
            
            buttons_layout.addWidget(add_button)
            buttons_layout.addWidget(remove_button)
            buttons_layout.addStretch()
            
            tab_layout.addWidget(self.list_widgets[tech])
            tab_layout.addLayout(buttons_layout)
            
            tab_widget.addTab(tab, tech)
        
        tech_layout.addWidget(tab_widget)
        tech_group.setLayout(tech_layout)
        main_layout.addWidget(tech_group)
        
        # Pulsanti OK e Annulla
        buttons_layout = QHBoxLayout()
        save_button = QPushButton("Salva")
        save_button.clicked.connect(self.save_config)
        
        cancel_button = QPushButton("Annulla")
        cancel_button.clicked.connect(self.reject)
        
        buttons_layout.addStretch()
        buttons_layout.addWidget(save_button)
        buttons_layout.addWidget(cancel_button)
        
        main_layout.addLayout(buttons_layout)
        
        # Carica la configurazione esistente
        self.load_config()

    def log_activity(self, message, type='info'):
        """Emette un segnale con il messaggio e il tipo"""
        #timestamp = datetime.datetime.now().strftime("[%H:%M:%S] ")
        #self.activityLogged.emit(timestamp + message, type)     
        self.activityLogged.emit(message, type)    
    
    def browse_directory(self):
        """Apre un dialogo per selezionare la directory di salvataggio"""
        directory = QFileDialog.getExistingDirectory(self, "Seleziona Directory", os.path.expanduser("~"))
        if directory:
            self.dir_input.setText(directory)
    
    def add_prefix(self, tech):
        """Aggiunge un nuovo prefisso alla lista della tecnologia specificata"""
        from PyQt5.QtWidgets import QInputDialog
        
        new_prefix, ok = QInputDialog.getText(self, f"Aggiungi Prefisso {tech}", 
                                             "Inserisci il nuovo prefisso:")
        
        if ok and new_prefix:
            # Verifica che il prefisso non sia già presente
            items = [self.list_widgets[tech].item(i).text() for i in range(self.list_widgets[tech].count())]
            
            if new_prefix in items:
                QMessageBox.warning(self, "Prefisso Duplicato", 
                                   f"Il prefisso '{new_prefix}' è già presente nella lista.")
                return
            
            self.list_widgets[tech].addItem(new_prefix)
            # Emetti segnale di Log dell'attività
            self.log_activity(f"Aggiunto prefisso '{new_prefix}' alla tecnologia {tech}", "info")
    
    def remove_prefix(self, tech):
        """Rimuove il prefisso selezionato dalla lista della tecnologia specificata"""
        selected_items = self.list_widgets[tech].selectedItems()
        
        if not selected_items:
            QMessageBox.warning(self, "Selezione Richiesta", 
                               "Seleziona un prefisso da rimuovere.")
            return
        
        for item in selected_items:
            prefix = item.text()
            row = self.list_widgets[tech].row(item)
            self.list_widgets[tech].takeItem(row)
            # Emetti segnale di Log dell'attività
            self.log_activity(f"Rimosso prefisso '{prefix}' dalla tecnologia {tech}", "info")
                
    
    def save_config(self):
        """Salva la configurazione in un file JSON"""
        # Carica la configurazione corrente dal file (se esiste)
        current_config = None
        if os.path.exists(self.config_file):
            try:
                with open(self.config_file, 'r') as f:
                    current_config = json.load(f)
            except Exception as e:
                # Se non riusciamo a leggere il file, assumiamo che sia diverso
                current_config = None
                print(f"Impossibile leggere la configurazione corrente: {str(e)}")
        
        # Crea la nuova configurazione
        new_config = {
            "save_directory": self.dir_input.text(),
            "technologies": {},
            "operations": {}
        }
        
        # Raccogli i prefissi per ogni tecnologia
        for tech, list_widget in self.list_widgets.items():
            prefixes = [list_widget.item(i).text() for i in range(list_widget.count())]
            new_config["technologies"][tech] = prefixes
        
        # Rileva lo stato delle checkbox
        new_config["operations"] = {
            "estrai_AdM": self.cb_estrai_adm.isChecked(),
            "estrai_OdM": self.cb_estrai_odm.isChecked(),
            "elabora_xls": self.cb_elabora_xls.isChecked()
        }

        # Confronta le configurazioni per vedere se ci sono state modifiche
        is_modified = True  # Default a True se non possiamo confrontare
        
        if current_config is not None:
            # Utilizziamo json.dumps per confrontare le rappresentazioni stringhe delle configurazioni
            # Con sort_keys=True, le chiavi saranno ordinate allo stesso modo
            current_str = json.dumps(current_config, sort_keys=True)
            new_str = json.dumps(new_config, sort_keys=True)
            is_modified = (current_str != new_str)
        
        # Se non ci sono modifiche, mostra un messaggio e termina
        if not is_modified:
            QMessageBox.information(self, "Nessuna Modifica", 
                                "Non sono state apportate modifiche alla configurazione.")
            # Emetti segnale di Log dell'attività
            self.log_activity("Non sono state apportate modifiche alla configurazione.", "info")
            self.accept()
            return

        # Salva la nuova configurazione nel file JSON
        try:
            with open(self.config_file, 'w') as f:
                json.dump(new_config, f, indent=4)
            
            QMessageBox.information(self, "Configurazione Salvata", 
                                "Le impostazioni sono state modificate e salvate con successo.")
            # Emetti segnale di Log dell'attività
            self.log_activity("Configurazione salvata con successo", "success")
            self.accept()
        except Exception as e:
            QMessageBox.critical(self, "Errore di Salvataggio", 
                                f"Impossibile salvare la configurazione: {str(e)}")
            # Emetti segnale di Log dell'attività
            self.log_activity("Impossibile salvare la configurazione", "error")
    
    def load_config(self):
        """Carica la configurazione dal file JSON se esiste"""
        if not os.path.exists(self.config_file):
            return
        
        try:
            with open(self.config_file, 'r') as f:
                config = json.load(f)
            
            # Imposta la directory di salvataggio
            if "save_directory" in config:
                self.dir_input.setText(config["save_directory"])
            
            # Imposta i prefissi per ogni tecnologia
            if "technologies" in config:
                for tech, prefixes in config["technologies"].items():
                    if tech in self.list_widgets:
                        # Rimuovi gli elementi predefiniti
                        self.list_widgets[tech].clear()
                        # Aggiungi i prefissi dalla configurazione
                        self.list_widgets[tech].addItems(prefixes)
            if "operations" in config:
                self.cb_estrai_adm.setChecked(config["operations"].get("estrai_AdM", False))
                self.cb_estrai_odm.setChecked(config["operations"].get("estrai_OdM", False))
                self.cb_elabora_xls.setChecked(config["operations"].get("elabora_xls", False))
        
        except Exception as e:
            QMessageBox.warning(self, "Errore di Caricamento", 
                               f"Impossibile caricare la configurazione: {str(e)}")

# Esempio di utilizzo:
# dialog = ConfigDialog()
# if dialog.exec_():
#     print("Configurazione salvata")