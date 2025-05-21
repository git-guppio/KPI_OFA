# kpi_ofa/ui/widgets/log_widget.py

import logging
from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QListWidget, QListWidgetItem, 
                            QMenu, QAction, QStyle, QApplication, QSizePolicy)
from PyQt5.QtGui import QCursor
from PyQt5.QtCore import Qt, pyqtSignal

logger = logging.getLogger(__name__)

class LogWidget(QWidget):
    """
    Widget per la visualizzazione e gestione dei log dell'applicazione.
    
    Fornisce funzionalità per visualizzare log con diverse icone in base al
    livello di importanza e un menu contestuale per copiare i messaggi.
    """
    
    # Segnale emesso quando viene registrato un messaggio di log
    logAdded = pyqtSignal(str, str)
    
    def __init__(self, parent=None):
        """
        Inizializza il widget di log.
        
        Args:
            parent: Widget genitore.
        """
        super().__init__(parent)
        
        # Layout principale
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        
        # Lista per i messaggi di log
        self.log_list = QListWidget()
        self.log_list.setMinimumWidth(200)
        self.log_list.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        
        # Attiva il menu contestuale
        self.log_list.setContextMenuPolicy(Qt.CustomContextMenu)
        self.log_list.customContextMenuRequested.connect(self.show_context_menu)
        
        layout.addWidget(self.log_list)
    
    def add_log(self, message, icon_type='info'):
        """
        Aggiunge un messaggio al log con un'icona appropriata.
        
        Args:
            message (str): Messaggio da aggiungere.
            icon_type (str): Tipo di icona ('info', 'error', 'success', 'warning', 'loading').
        """
        item = QListWidgetItem(message)
        
        # Ottieni l'istanza di style dal widget
        style = self.style()
        
        # Imposta l'icona in base al tipo
        if icon_type == 'info':
            item.setIcon(style.standardIcon(QStyle.SP_MessageBoxInformation))
        elif icon_type == 'error' or icon_type == 'critical':
            item.setIcon(style.standardIcon(QStyle.SP_MessageBoxCritical))
        elif icon_type == 'success':
            item.setIcon(style.standardIcon(QStyle.SP_DialogApplyButton))
        elif icon_type == 'warning':
            item.setIcon(style.standardIcon(QStyle.SP_MessageBoxWarning))
        elif icon_type == 'loading':
            item.setIcon(style.standardIcon(QStyle.SP_BrowserReload))
        
        self.log_list.addItem(item)
        self.log_list.scrollToBottom()
        
        # Emetti il segnale con il messaggio e il tipo di icona
        self.logAdded.emit(message, icon_type)
        
        # Logga anche con il logger standard
        log_level_map = {
            'info': logging.INFO,
            'error': logging.ERROR,
            'critical': logging.CRITICAL,
            'warning': logging.WARNING,
            'success': logging.INFO,
            'loading': logging.INFO
        }
        
        log_level = log_level_map.get(icon_type, logging.INFO)
        logger.log(log_level, message)
    
    def clear_logs(self):
        """Pulisce tutti i messaggi di log dalla lista."""
        self.log_list.clear()
        logger.info("Log cancellati")
    
    def show_context_menu(self, position):
        """
        Mostra il menu contestuale con opzioni per copiare i log.
        
        Args:
            position: Posizione del cursore.
        """
        # Crea menu contestuale
        context_menu = QMenu(self)
        
        # Aggiungi l'azione "Copia"
        copy_action = QAction("Copia elemento", self)
        copy_action.triggered.connect(self.copy_selected_items)
        context_menu.addAction(copy_action)
        
        # Aggiungi l'azione "Copia tutto"
        copy_all_action = QAction("Copia tutto", self)
        copy_all_action.triggered.connect(self.copy_all_items)
        context_menu.addAction(copy_all_action)
        
        # Aggiungi separatore
        context_menu.addSeparator()
        
        # Aggiungi l'azione "Cancella log"
        clear_action = QAction("Cancella log", self)
        clear_action.triggered.connect(self.clear_logs)
        context_menu.addAction(clear_action)
        
        # Mostra il menu alla posizione del cursore
        context_menu.exec_(QCursor.pos())
    
    def copy_selected_items(self):
        """Copia gli elementi selezionati negli appunti."""
        selected_items = self.log_list.selectedItems()
        if selected_items:
            text = "\n".join(item.text() for item in selected_items)
            QApplication.clipboard().setText(text)
            logger.debug("Elementi selezionati copiati negli appunti")
    
    def copy_all_items(self):
        """Copia tutti gli elementi negli appunti."""
        all_items = []
        for i in range(self.log_list.count()):
            all_items.append(self.log_list.item(i).text())
        
        text = "\n".join(all_items)
        QApplication.clipboard().setText(text)
        logger.debug("Tutti gli elementi copiati negli appunti")
    
    def get_log_count(self):
        """
        Restituisce il numero di messaggi di log.
        
        Returns:
            int: Numero di messaggi nel log.
        """
        return self.log_list.count()
    
    def get_log_text(self):
        """
        Restituisce il testo di tutti i messaggi di log.
        
        Returns:
            str: Testo di tutti i messaggi.
        """
        return "\n".join(self.log_list.item(i).text() for i in range(self.log_list.count()))