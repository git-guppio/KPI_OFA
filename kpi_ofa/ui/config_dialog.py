# kpi_ofa/ui/config_dialog.py

import os
import json
import logging
from PyQt5.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QLabel, 
                            QLineEdit, QPushButton, QFileDialog, QTabWidget,
                            QListWidget, QGroupBox, QMessageBox, QWidget,
                            QCheckBox, QInputDialog, QListWidgetItem)
from PyQt5.QtCore import Qt, pyqtSignal

logger = logging.getLogger(__name__)

class ConfigDialog(QDialog):
    """
    Finestra di dialogo per la configurazione dell'applicazione.
    
    Permette di configurare:
    - La directory di salvataggio dei dati
    - Le operazioni da eseguire
    - I prefissi delle tecnologie
    """
    
    # Segnale per il log delle attività
    activityLogged = pyqtSignal(str, str)  # (message, icon_type)
    
    def __init__(self, config_manager, parent=None):
        """
        Inizializza il dialogo di configurazione.
        
        Args:
            config_manager: Gestore della configurazione.
            parent: Widget genitore.
        """
        super().__init__(parent)
        
        # Salva il riferimento al gestore della configurazione
        self.config_manager = config_manager
        
        # Ottieni la configurazione corrente
        self.config = config_manager.get_config()
        
        # Configurazione del dialogo
        self.setWindowTitle("Configurazione")
        self.setMinimumSize(500, 400)
        
        # Crea l'interfaccia utente
        self.setup_ui()
        
        # Carica la configurazione nei controlli
        self.load_config_to_ui()
        
        # Log dell'inizializzazione
        logger.info("Finestra di configurazione inizializzata")
    
    def setup_ui(self):
        """Crea l'interfaccia utente del dialogo di configurazione."""
        # Layout principale
        main_layout = QVBoxLayout(self)
        
        # Sezione directory
        self.create_directory_section(main_layout)
        
        # Sezione operazioni
        self.create_operations_section(main_layout)
        
        # Sezione tecnologie
        self.create_technologies_section(main_layout)
        
        # Pulsanti OK e Annulla
        self.create_buttons_section(main_layout)
    
    def create_directory_section(self, parent_layout):
        """
        Crea la sezione per la configurazione della directory di salvataggio.
        
        Args:
            parent_layout: Layout genitore a cui aggiungere la sezione.
        """
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
        parent_layout.addWidget(dir_group)
    
    def create_operations_section(self, parent_layout):
        """
        Crea la sezione per la configurazione delle operazioni da eseguire.
        
        Args:
            parent_layout: Layout genitore a cui aggiungere la sezione.
        """
        op_group = QGroupBox("Operazioni da eseguire")
        op_layout = QHBoxLayout()
        
        # Crea le checkbox
        self.cb_estrai_adm = QCheckBox("Estrai AdM")
        self.cb_estrai_odm = QCheckBox("Estrai OdM")
        self.cb_estrai_AFKO = QCheckBox("Estrai AFKO")
        self.cb_elabora_xls = QCheckBox("Elabora File XLS")
        
        # Aggiungi tooltip per spiegare le operazioni
        self.cb_estrai_adm.setToolTip("Estrai Avvisi di Manutenzione (AdM) da SAP")
        self.cb_estrai_odm.setToolTip("Estrai Ordini di Manutenzione (OdM) da SAP")
        self.cb_estrai_AFKO.setToolTip("Estrai Tabella AFKO tramite transazione SE16 da SAP")
        self.cb_elabora_xls.setToolTip("Elabora i dati estratti in un file Excel")
        
        # Aggiungi le checkbox al layout
        op_layout.addWidget(self.cb_estrai_adm)
        op_layout.addWidget(self.cb_estrai_odm)
        op_layout.addWidget(self.cb_estrai_AFKO)
        op_layout.addWidget(self.cb_elabora_xls)
        
        # Imposta il layout del gruppo
        op_group.setLayout(op_layout)
        parent_layout.addWidget(op_group)
    
    def create_technologies_section(self, parent_layout):
        """
        Crea la sezione per la configurazione delle tecnologie.
        
        Args:
            parent_layout: Layout genitore a cui aggiungere la sezione.
        """
        tech_group = QGroupBox("Configurazione Tecnologie")
        tech_layout = QVBoxLayout()
        
        # Crea un widget con tab per le diverse tecnologie
        self.tab_widget = QTabWidget()
        
        # Definizione delle tecnologie di default e relativi prefissi
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
            
            # Liste dei prefissi
            list_widget = QListWidget()
            list_widget.setSelectionMode(QListWidget.ExtendedSelection)
            self.list_widgets[tech] = list_widget
            
            # Pulsanti per aggiungere/rimuovere prefissi
            buttons_layout = QHBoxLayout()
            add_button = QPushButton("Aggiungi")
            add_button.clicked.connect(lambda checked, t=tech: self.add_prefix(t))
            
            remove_button = QPushButton("Rimuovi")
            remove_button.clicked.connect(lambda checked, t=tech: self.remove_prefix(t))
            
            buttons_layout.addWidget(add_button)
            buttons_layout.addWidget(remove_button)
            buttons_layout.addStretch()
            
            # Aggiungi i widget al layout della tab
            tab_layout.addWidget(list_widget)
            tab_layout.addLayout(buttons_layout)
            
            # Aggiungi la tab al widget tabs
            self.tab_widget.addTab(tab, tech)
        
        # Aggiungi il widget tabs al layout delle tecnologie
        tech_layout.addWidget(self.tab_widget)
        tech_group.setLayout(tech_layout)
        parent_layout.addWidget(tech_group)
    
    def create_buttons_section(self, parent_layout):
        """
        Crea la sezione dei pulsanti OK e Annulla.
        
        Args:
            parent_layout: Layout genitore a cui aggiungere la sezione.
        """
        buttons_layout = QHBoxLayout()
        
        save_button = QPushButton("Salva")
        save_button.clicked.connect(self.save_config)
        save_button.setDefault(True)
        
        cancel_button = QPushButton("Annulla")
        cancel_button.clicked.connect(self.reject)
        
        buttons_layout.addStretch()
        buttons_layout.addWidget(save_button)
        buttons_layout.addWidget(cancel_button)
        
        parent_layout.addLayout(buttons_layout)
    
    def log_activity(self, message, level='info'):
        """
        Emette un segnale con il messaggio e il tipo di log.
        
        Args:
            message (str): Messaggio da registrare.
            level (str): Livello del log ('info', 'warning', 'error', 'success').
        """
        # Emetti il segnale per la UI
        self.activityLogged.emit(message, level)
        
        # Logga anche con il logger standard
        log_method = getattr(logger, level) if hasattr(logger, level) else logger.info
        log_method(message)
    
    def load_config_to_ui(self):
        """Carica la configurazione esistente nei controlli dell'interfaccia."""
        # Imposta la directory di salvataggio
        save_directory = self.config.get("save_directory", "")
        if save_directory:
            self.dir_input.setText(save_directory)
        
        # Imposta lo stato delle checkbox per le operazioni
        operations = self.config.get("operations", {})
        self.cb_estrai_adm.setChecked(operations.get("estrai_AdM", False))
        self.cb_estrai_odm.setChecked(operations.get("estrai_OdM", False))
        self.cb_estrai_AFKO.setChecked(operations.get("estrai_AFKO", False))
        self.cb_elabora_xls.setChecked(operations.get("elabora_xls", False))
        
        # Imposta i prefissi per ogni tecnologia
        tech_config = self.config.get("technologies", {})
        for tech, prefixes in tech_config.items():
            if tech in self.list_widgets:
                list_widget = self.list_widgets[tech]
                
                # Rimuovi gli elementi predefiniti
                list_widget.clear()
                
                # Aggiungi i prefissi dalla configurazione
                for prefix in prefixes:
                    item = QListWidgetItem(prefix)
                    list_widget.addItem(item)
        
        self.log_activity("Configurazione caricata nei controlli")
    
    def browse_directory(self):
        """Apre un dialogo per selezionare la directory di salvataggio."""
        current_dir = self.dir_input.text() or os.path.expanduser("~")
        directory = QFileDialog.getExistingDirectory(
            self, 
            "Seleziona Directory", 
            current_dir
        )
        
        if directory:
            self.dir_input.setText(directory)
            self.log_activity(f"Selezionata directory: {directory}")
    
    def add_prefix(self, tech):
        """
        Aggiunge un nuovo prefisso alla tecnologia specificata.
        
        Args:
            tech (str): Nome della tecnologia.
        """
        new_prefix, ok = QInputDialog.getText(
            self, 
            f"Aggiungi Prefisso {tech}", 
            "Inserisci il nuovo prefisso:"
        )
        
        if ok and new_prefix:
            # Verifica che il prefisso non sia già presente
            list_widget = self.list_widgets[tech]
            items = [list_widget.item(i).text() for i in range(list_widget.count())]
            
            if new_prefix in items:
                QMessageBox.warning(
                    self, 
                    "Prefisso Duplicato", 
                    f"Il prefisso '{new_prefix}' è già presente nella lista."
                )
                return
            
            # Aggiungi il nuovo prefisso alla lista
            list_widget.addItem(new_prefix)
            self.log_activity(f"Aggiunto prefisso '{new_prefix}' alla tecnologia {tech}")
    
    def remove_prefix(self, tech):
        """
        Rimuove i prefissi selezionati dalla tecnologia specificata.
        
        Args:
            tech (str): Nome della tecnologia.
        """
        list_widget = self.list_widgets[tech]
        selected_items = list_widget.selectedItems()
        
        if not selected_items:
            QMessageBox.warning(
                self, 
                "Selezione Richiesta", 
                "Seleziona uno o più prefissi da rimuovere."
            )
            return
        
        # Conferma la rimozione
        num_items = len(selected_items)
        confirm_message = (f"Sei sicuro di voler rimuovere {num_items} "
                           f"prefiss{'o' if num_items == 1 else 'i'} dalla tecnologia {tech}?")
        
        reply = QMessageBox.question(
            self, 
            "Conferma Rimozione", 
            confirm_message,
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )
        
        if reply == QMessageBox.Yes:
            # Rimuovi gli elementi selezionati in ordine inverso per evitare problemi di indice
            rows_to_remove = [list_widget.row(item) for item in selected_items]
            rows_to_remove.sort(reverse=True)
            
            for row in rows_to_remove:
                item = list_widget.takeItem(row)
                self.log_activity(f"Rimosso prefisso '{item.text()}' dalla tecnologia {tech}")
    
    def get_config_from_ui(self):
        """
        Raccoglie la configurazione dall'interfaccia utente.
        
        Returns:
            dict: Configurazione raccolta dall'interfaccia.
        """
        new_config = {
            "save_directory": self.dir_input.text(),
            "technologies": {},
            "operations": {}
        }
        
        # Raccogli i prefissi per ogni tecnologia
        for tech, list_widget in self.list_widgets.items():
            prefixes = [list_widget.item(i).text() for i in range(list_widget.count())]
            new_config["technologies"][tech] = prefixes
        
        # Raccogli lo stato delle checkbox
        new_config["operations"] = {
            "estrai_AdM": self.cb_estrai_adm.isChecked(),
            "estrai_OdM": self.cb_estrai_odm.isChecked(),
            "estrai_AFKO": self.cb_estrai_AFKO.isChecked(),
            "elabora_xls": self.cb_elabora_xls.isChecked()
        }
        
        return new_config
    
    def validate_config(self, config):
        """
        Valida la configurazione raccolta dall'interfaccia.
        
        Args:
            config (dict): Configurazione da validare.
            
        Returns:
            tuple: (bool, str) - True se la configurazione è valida, False altrimenti,
                  e un messaggio di errore in caso di validazione fallita.
        """
        # Verifica che la directory di salvataggio sia specificata
        save_dir = config.get("save_directory", "")
        if not save_dir:
            return False, "La directory di salvataggio non è stata specificata."
        
        # Verifica che la directory esista
        if not os.path.exists(save_dir):
            return False, f"La directory '{save_dir}' non esiste."
        
        # Verifica che ci sia almeno un'operazione attiva
        operations = config.get("operations", {})
        if not any(operations.values()):
            return False, "È necessario selezionare almeno un'operazione da eseguire."
        
        # Verifica che ci sia almeno una tecnologia con prefissi
        technologies = config.get("technologies", {})
        has_prefixes = False
        for tech, prefixes in technologies.items():
            if prefixes:
                has_prefixes = True
                break
        
        if not has_prefixes:
            return False, "È necessario specificare almeno un prefisso per una tecnologia."
        
        # Tutte le verifiche sono state superate
        return True, ""
    
    def save_config(self):
        """Salva la configurazione raccolta dall'interfaccia."""
        # Ottieni la configurazione dall'interfaccia
        new_config = self.get_config_from_ui()
        
        # Valida la configurazione
        is_valid, error_message = self.validate_config(new_config)
        if not is_valid:
            QMessageBox.warning(self, "Configurazione non valida", error_message)
            self.log_activity(f"Validazione configurazione fallita: {error_message}", "error")
            return
        
        # Confronta la nuova configurazione con quella attuale
        current_config_str = json.dumps(self.config, sort_keys=True)
        new_config_str = json.dumps(new_config, sort_keys=True)
        
        # Se non ci sono modifiche, non salvare
        if current_config_str == new_config_str:
            QMessageBox.information(
                self, 
                "Nessuna Modifica", 
                "Non sono state apportate modifiche alla configurazione."
            )
            self.log_activity("Non sono state apportate modifiche alla configurazione")
            self.accept()
            return
        
        # Salva la configurazione
        if self.config_manager.save_config(new_config):
            QMessageBox.information(
                self, 
                "Configurazione Salvata", 
                "Le impostazioni sono state modificate e salvate con successo."
            )
            self.log_activity("Configurazione salvata con successo", "success")
            self.accept()
        else:
            QMessageBox.critical(
                self, 
                "Errore di Salvataggio", 
                "Impossibile salvare la configurazione."
            )
            self.log_activity("Impossibile salvare la configurazione", "error")
    
    def closeEvent(self, event):
        """
        Gestisce l'evento di chiusura della finestra.
        
        Args:
            event: Evento di chiusura.
        """
        # Se ci sono modifiche non salvate, chiedi conferma
        new_config = self.get_config_from_ui()
        current_config_str = json.dumps(self.config, sort_keys=True)
        new_config_str = json.dumps(new_config, sort_keys=True)
        
        if current_config_str != new_config_str:
            reply = QMessageBox.question(
                self,
                "Modifiche non salvate",
                "Ci sono modifiche non salvate. Vuoi salvare prima di uscire?",
                QMessageBox.Save | QMessageBox.Discard | QMessageBox.Cancel,
                QMessageBox.Save
            )
            
            if reply == QMessageBox.Save:
                self.save_config()
                event.accept()
            elif reply == QMessageBox.Discard:
                self.log_activity("Modifiche alla configurazione scartate", "warning")
                event.accept()
            else:
                event.ignore()
        else:
            event.accept()