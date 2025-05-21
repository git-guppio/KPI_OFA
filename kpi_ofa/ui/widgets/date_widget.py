# kpi_ofa/ui/widgets/date_widget.py

import logging
from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QLabel, 
                            QDateEdit, QMessageBox)
from PyQt5.QtCore import QDate, Qt, pyqtSignal

logger = logging.getLogger(__name__)

class DateRangeWidget(QWidget):
    """
    Widget per la selezione di un intervallo di date.
    
    Fornisce controlli per selezionare data di inizio e fine, con
    validazione dell'intervallo e segnali per notificare i cambiamenti.
    """
    
    # Segnale emesso quando le date cambiano
    dateRangeChanged = pyqtSignal(QDate, QDate, bool)
    
    def __init__(self, parent=None):
        """
        Inizializza il widget per l'intervallo di date.
        
        Args:
            parent: Widget genitore.
        """
        super().__init__(parent)
        
        # Layout principale
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(0, 0, 0, 0)
        
        # Riga 1 - Data inizio
        row1_layout = QHBoxLayout()
        start_date_label = QLabel("Data inizio:")
        start_date_label.setFixedWidth(60)
        self.start_date_picker = QDateEdit()
        self.start_date_picker.setDate(QDate.currentDate())
        self.start_date_picker.setCalendarPopup(True)
        self.start_date_picker.setMinimumWidth(100)
        self.start_date_picker.setDisplayFormat("dd/MM/yyyy")
        row1_layout.addWidget(start_date_label)
        row1_layout.addWidget(self.start_date_picker)
        row1_layout.addStretch()
        
        # Riga 2 - Data fine
        row2_layout = QHBoxLayout()
        end_date_label = QLabel("Data fine:")
        end_date_label.setFixedWidth(60)
        self.end_date_picker = QDateEdit()
        self.end_date_picker.setDate(QDate.currentDate().addDays(7))  # Una settimana in avanti di default
        self.end_date_picker.setCalendarPopup(True)
        self.end_date_picker.setMinimumWidth(100)
        self.end_date_picker.setDisplayFormat("dd/MM/yyyy")
        row2_layout.addWidget(end_date_label)
        row2_layout.addWidget(self.end_date_picker)
        row2_layout.addStretch()
        
        # Aggiungi le righe al layout principale
        main_layout.addLayout(row1_layout)
        main_layout.addLayout(row2_layout)
        
        # Connetti i segnali
        self.start_date_picker.dateChanged.connect(self.on_date_changed)
        self.end_date_picker.dateChanged.connect(self.on_date_changed)
        
        # Emetti il segnale iniziale
        self.on_date_changed()
        
        logger.debug("DateRangeWidget inizializzato")
    
    def on_date_changed(self):
        """
        Gestisce il cambio di data e emette il segnale dateRangeChanged.
        
        Verifica la validità dell'intervallo selezionato e emette un segnale
        con le date e un flag che indica se l'intervallo è valido.
        """
        start_date = self.start_date_picker.date()
        end_date = self.end_date_picker.date()
        
        # Verifica la validità dell'intervallo
        is_valid = start_date <= end_date
        
        # Emetti il segnale con le date e la validità
        self.dateRangeChanged.emit(start_date, end_date, is_valid)
        
        # Log dell'evento
        if is_valid:
            logger.debug(f"Intervallo date cambiato: {start_date.toString('dd/MM/yyyy')} - {end_date.toString('dd/MM/yyyy')}")
        else:
            logger.warning(f"Intervallo date non valido: {start_date.toString('dd/MM/yyyy')} - {end_date.toString('dd/MM/yyyy')}")
    
    def get_date_range(self):
        """
        Restituisce l'intervallo di date corrente.
        
        Returns:
            tuple: (QDate, QDate) - Data di inizio e fine.
        """
        return (self.start_date_picker.date(), self.end_date_picker.date())
    
    def set_date_range(self, start_date, end_date):
        """
        Imposta l'intervallo di date.
        
        Args:
            start_date (QDate): Data di inizio.
            end_date (QDate): Data di fine.
        """
        # Blocca temporaneamente i segnali per evitare emissioni multiple
        old_start_block = self.start_date_picker.blockSignals(True)
        old_end_block = self.end_date_picker.blockSignals(True)
        
        self.start_date_picker.setDate(start_date)
        self.end_date_picker.setDate(end_date)
        
        # Ripristina il blocco segnali
        self.start_date_picker.blockSignals(old_start_block)
        self.end_date_picker.blockSignals(old_end_block)
        
        # Emetti manualmente il segnale
        self.on_date_changed()
        
        logger.debug(f"Intervallo date impostato: {start_date.toString('dd/MM/yyyy')} - {end_date.toString('dd/MM/yyyy')}")
    
    def validate_date_range(self):
        """
        Verifica la validità dell'intervallo di date e mostra messaggi di errore se necessario.
        
        Returns:
            bool: True se l'intervallo è valido, False altrimenti.
        """
        start_date, end_date = self.get_date_range()
        
        # Verifica che la data di inizio sia precedente o uguale alla data di fine
        if start_date > end_date:
            error_message = "La data di inizio non può essere successiva alla data di fine"
            logger.error(error_message)
            QMessageBox.warning(self, "Errore nell'intervallo date", error_message)
            return False
        
        # Verifica che la data di inizio non sia uguale alla data di fine
        if start_date == end_date:
            error_message = "La data di inizio non può essere uguale alla data di fine"
            logger.error(error_message)
            QMessageBox.warning(self, "Errore nell'intervallo date", error_message)
            return False
        
        # Calcola la differenza in giorni
        days_difference = start_date.daysTo(end_date)
        
        # Verifica che l'intervallo non superi un anno (366 giorni per considerare anni bisestili)
        max_days = 366
        if days_difference > max_days:
            error_message = f"L'intervallo di date non può superare un anno ({max_days} giorni)"
            logger.error(error_message)
            QMessageBox.warning(self, "Intervallo troppo ampio", error_message)
            return False
        
        # Tutte le verifiche sono state superate
        logger.info(f"Intervallo di date valido: {start_date.toString('dd/MM/yyyy')} - {end_date.toString('dd/MM/yyyy')}")
        return True
    
    def set_default_range(self, days=7):
        """
        Imposta l'intervallo di date ai valori predefiniti.
        
        Args:
            days (int): Numero di giorni da aggiungere alla data corrente.
        """
        current_date = QDate.currentDate()
        self.set_date_range(current_date, current_date.addDays(days))
        
        logger.info(f"Intervallo date reimpostato ai valori predefiniti: oggi + {days} giorni")
    
    def set_min_max_dates(self, min_date=None, max_date=None):
        """
        Imposta le date minime e massime selezionabili.
        
        Args:
            min_date (QDate, optional): Data minima selezionabile.
            max_date (QDate, optional): Data massima selezionabile.
        """
        if min_date:
            self.start_date_picker.setMinimumDate(min_date)
            self.end_date_picker.setMinimumDate(min_date)
        
        if max_date:
            self.start_date_picker.setMaximumDate(max_date)
            self.end_date_picker.setMaximumDate(max_date)
        
        logger.debug(f"Limiti date impostati: min={min_date.toString('dd/MM/yyyy') if min_date else 'None'}, "
                     f"max={max_date.toString('dd/MM/yyyy') if max_date else 'None'}")