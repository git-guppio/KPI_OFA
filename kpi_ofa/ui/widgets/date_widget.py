# kpi_ofa/ui/widgets/date_widget.py

import logging
from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QLabel,
                            QDateEdit, QMessageBox, QCalendarWidget)
from PyQt5.QtCore import QDate, Qt, pyqtSignal, QEvent

logger = logging.getLogger(__name__)

class DateRangeWidget(QWidget):
    """
    Widget per la selezione di un intervallo di date.

    Fornisce controlli per selezionare data di inizio e fine, con
    validazione dell'intervallo e segnali per notificare i cambiamenti.
    """

    # Data sentinella per indicare "nessuna data selezionata"
    SENTINEL_DATE = QDate(1900, 1, 1)

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
        self.start_date_picker.setCalendarPopup(True)
        self.start_date_picker.setMinimumWidth(100)
        self.start_date_picker.setDisplayFormat("dd/MM/yyyy")
        # Configura data sentinella con placeholder
        self.start_date_picker.setMinimumDate(self.SENTINEL_DATE)
        self.start_date_picker.setSpecialValueText("gg/mm/aaaa")
        self.start_date_picker.setDate(self.SENTINEL_DATE)
        row1_layout.addWidget(start_date_label)
        row1_layout.addWidget(self.start_date_picker)
        row1_layout.addStretch()

        # Riga 2 - Data fine
        row2_layout = QHBoxLayout()
        end_date_label = QLabel("Data fine:")
        end_date_label.setFixedWidth(60)
        self.end_date_picker = QDateEdit()
        self.end_date_picker.setCalendarPopup(True)
        self.end_date_picker.setMinimumWidth(100)
        self.end_date_picker.setDisplayFormat("dd/MM/yyyy")
        # Configura data sentinella con placeholder
        self.end_date_picker.setMinimumDate(self.SENTINEL_DATE)
        self.end_date_picker.setSpecialValueText("gg/mm/aaaa")
        self.end_date_picker.setDate(self.SENTINEL_DATE)
        row2_layout.addWidget(end_date_label)
        row2_layout.addWidget(self.end_date_picker)
        row2_layout.addStretch()

        # Aggiungi le righe al layout principale
        main_layout.addLayout(row1_layout)
        main_layout.addLayout(row2_layout)

        # Connetti i segnali
        self.start_date_picker.dateChanged.connect(self.on_date_changed)
        self.end_date_picker.dateChanged.connect(self.on_date_changed)

        # Installa event filter sui calendari per navigare alla data odierna all'apertura
        self.start_date_picker.calendarWidget().installEventFilter(self)
        self.end_date_picker.calendarWidget().installEventFilter(self)

        # Emetti il segnale iniziale
        self.on_date_changed()

        logger.debug("DateRangeWidget inizializzato")

    def eventFilter(self, obj, event):
        """
        Filtra gli eventi per navigare alla data odierna quando il calendario viene aperto.

        Quando la data impostata è la sentinella (1900-01-01), il calendario viene
        automaticamente posizionato sul mese corrente per facilitare la selezione.

        Args:
            obj: Oggetto che ha generato l'evento.
            event: Evento da filtrare.

        Returns:
            bool: False per permettere la propagazione dell'evento.
        """
        if event.type() == QEvent.Show and isinstance(obj, QCalendarWidget):
            # Identifica quale date picker ha aperto il calendario
            current_date = None
            if obj == self.start_date_picker.calendarWidget():
                current_date = self.start_date_picker.date()
            elif obj == self.end_date_picker.calendarWidget():
                current_date = self.end_date_picker.date()

            # Se la data è la sentinella, naviga al mese corrente
            if current_date is not None and current_date == self.SENTINEL_DATE:
                today = QDate.currentDate()
                obj.setCurrentPage(today.year(), today.month())

        return super().eventFilter(obj, event)
    
    def on_date_changed(self):
        """
        Gestisce il cambio di data e emette il segnale dateRangeChanged.

        Verifica la validità dell'intervallo selezionato e emette un segnale
        con le date e un flag che indica se l'intervallo è valido.
        """
        start_date = self.start_date_picker.date()
        end_date = self.end_date_picker.date()

        # Verifica la validità completa dell'intervallo
        is_valid = self.are_dates_valid()

        # Emetti il segnale con le date e la validità
        self.dateRangeChanged.emit(start_date, end_date, is_valid)

        # Log dell'evento
        if self.are_dates_set():
            if is_valid:
                logger.debug(f"Intervallo date cambiato: {start_date.toString('dd/MM/yyyy')} - {end_date.toString('dd/MM/yyyy')}")
            else:
                logger.warning(f"Intervallo date non valido: {start_date.toString('dd/MM/yyyy')} - {end_date.toString('dd/MM/yyyy')}")
        else:
            logger.debug("Date non ancora selezionate (placeholder attivo)")

    def are_dates_set(self):
        """
        Verifica se entrambe le date sono state impostate (non sono sentinelle).

        Returns:
            bool: True se entrambe le date sono state selezionate, False altrimenti.
        """
        start_date = self.start_date_picker.date()
        end_date = self.end_date_picker.date()
        return start_date != self.SENTINEL_DATE and end_date != self.SENTINEL_DATE

    def are_dates_valid(self):
        """
        Verifica la validità completa delle date selezionate.

        Condizioni verificate:
        - Entrambe le date devono essere impostate (non sentinelle)
        - La data di inizio deve essere precedente alla data di fine
        - Il range massimo è di 2 mesi (62 giorni)
        - Entrambe le date devono essere nel passato

        Returns:
            bool: True se tutte le condizioni sono soddisfatte, False altrimenti.
        """
        # Verifica che le date siano state impostate
        if not self.are_dates_set():
            return False

        start_date = self.start_date_picker.date()
        end_date = self.end_date_picker.date()
        today = QDate.currentDate()

        # Verifica che la data di inizio sia precedente alla data di fine
        if start_date >= end_date:
            return False

        # Verifica che il range non superi 2 mesi (circa 62 giorni)
        max_days = 62
        if start_date.daysTo(end_date) > max_days:
            return False

        # Verifica che entrambe le date siano nel passato
        if start_date >= today or end_date > today:
            return False

        return True

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

        Condizioni verificate:
        - Entrambe le date devono essere impostate
        - La data di inizio deve essere precedente alla data di fine
        - Il range massimo è di 2 mesi (62 giorni)
        - Entrambe le date devono essere nel passato

        Returns:
            bool: True se l'intervallo è valido, False altrimenti.
        """
        start_date, end_date = self.get_date_range()
        today = QDate.currentDate()

        # Verifica che le date siano state impostate
        if not self.are_dates_set():
            error_message = "Selezionare entrambe le date prima di procedere"
            logger.error(error_message)
            QMessageBox.warning(self, "Date non selezionate", error_message)
            return False

        # Verifica che la data di inizio sia precedente alla data di fine
        if start_date >= end_date:
            error_message = "La data di inizio deve essere precedente alla data di fine"
            logger.error(error_message)
            QMessageBox.warning(self, "Errore nell'intervallo date", error_message)
            return False

        # Calcola la differenza in giorni
        days_difference = start_date.daysTo(end_date)

        # Verifica che l'intervallo non superi 2 mesi (62 giorni)
        max_days = 62
        if days_difference > max_days:
            error_message = f"L'intervallo di date non può superare 2 mesi ({max_days} giorni)"
            logger.error(error_message)
            QMessageBox.warning(self, "Intervallo troppo ampio", error_message)
            return False

        # Verifica che entrambe le date siano nel passato
        if start_date >= today:
            error_message = "La data di inizio deve essere nel passato"
            logger.error(error_message)
            QMessageBox.warning(self, "Data non valida", error_message)
            return False

        if end_date > today:
            error_message = "La data di fine non può essere nel futuro"
            logger.error(error_message)
            QMessageBox.warning(self, "Data non valida", error_message)
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

    def reset_to_empty(self):
        """
        Resetta le date al valore sentinella, mostrando il placeholder.

        Questo metodo viene utilizzato per resettare i campi data allo stato
        iniziale "non selezionato", visualizzando il testo 'gg/mm/aaaa'.
        """
        self.set_date_range(self.SENTINEL_DATE, self.SENTINEL_DATE)
        logger.info("Date resettate al valore sentinella (placeholder attivo)")
    
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