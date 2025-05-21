# kpi_ofa/core/log_manager.py
import logging
import time
from PyQt5.QtWidgets import QApplication

class LogManager:
    """
    Gestore centralizzato per il logging dell'applicazione.
    Implementa il pattern Singleton per garantire un'unica istanza.

    Nota bene: In ogni classe che utilizza questo gestore, è necessario
    importare il modulo logging e configurarlo correttamente.
    Esempio:
        import logging
        from kpi_ofa.core.log_manager import LogManager
        logging.basicConfig(level=logging.INFO)
        logger = logging.getLogger("kpi_ofa")

    Utilizzo:
        log_manager = LogManager()
        log_manager.log("Messaggio di log", level="info", origin=logger.name)
    """     
    _instance = None
    
    def __new__(cls):
        if cls._instance is None:
            cls._instance = super(LogManager, cls).__new__(cls)
            cls._instance._initialized = False
        return cls._instance
    
    def __init__(self):
        """Inizializza il gestore dei log (solo alla prima istanziazione)"""
        if self._initialized:
            return
            
        # Setup del logger
        self.logger = logging.getLogger("kpi_ofa")
        
        # Attributi per la gestione dello stato
        self._last_status_level = ""
        self._last_status_time = 0
        self._last_status_min_time = 0
        
        # Callback per UI (inizializzato a None)
        self.ui_log_callback = None # Callback per il log nell'interfaccia utente (widget di log)
        self.status_bar_callback = None # Callback per la barra di stato (statusBar)
        
        self._initialized = True
    
    def set_ui_callback(self, callback):
        """Imposta il callback per il log nell'interfaccia utente"""
        self.ui_log_callback = callback
    
    def set_status_bar_callback(self, callback):
        """Imposta il callback per l'aggiornamento della barra di stato"""
        self.status_bar_callback = callback
    
    def log(self, message, level="info", update_status=True, update_log=True, 
            min_display_seconds=None, origin="UnKnown", *args, **kwargs):
        """
        Metodo unificato per gestire log attraverso tutti i canali disponibili
        
        Args:
            message: Il messaggio da registrare
            level: Livello del log ('info', 'warning', 'error', ecc.)
            update_status: Se aggiornare la statusBar
            update_log: Se aggiornare il widget di log visuale
            min_display_seconds: Tempo minimo di visualizzazione
            origin: Origine del messaggio per evitare duplicazioni può essere utilizzato origin=logger.name
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
        
        # 1. Logga con il logger standard di Python SOLO se l'origine è 'main'
        # Se origin non è 'main', significa che il messaggio proviene già da un logger specifico
        # quindi non dobbiamo registrarlo di nuovo nel logger principale
        if origin == "UnKnown":  # Controlla se l'origine è sconosciuta
            # Mappa i livelli speciali alle funzioni del logger
            logger_level_map = {
                "success": "info",    # success non è un livello standard, usa info
                "loading": "info"     # loading non è un livello standard, usa info
            }
            # Ottieni il livello standard corrispondente o usa lo stesso se è già standard
            logger_level = logger_level_map.get(level, level)
            logger_method = getattr(self.logger, logger_level) if hasattr(self.logger, logger_level) else self.logger.info
            logger_method(formatted_message)
        
        # 2. Invia il log all'interfaccia utente se il callback è impostato
        if update_log and self.ui_log_callback:
            self.ui_log_callback(formatted_message, level)
        
        # 3. Aggiorna la barra di stato se richiesto e se il callback è impostato
        if update_status and self.status_bar_callback:
            self._update_status_bar(formatted_message, level, min_display_seconds)

    def _update_status_bar(self, message, level, min_display_seconds):
        """
        Gestisce l'aggiornamento della barra di stato
        
        Args:
            message: Messaggio da visualizzare
            level: Livello del messaggio
            min_display_seconds: Tempo minimo di visualizzazione
        """
        # Livelli in ordine crescente di importanza
        level_importance = ["debug", "info", "success", "loading", "warning", "error", "critical"]
        
        # Tempi minimi di visualizzazione predefiniti per livello (in secondi)
        default_min_times = {
            "debug": 2, "info": 3, "success": 5, "loading": 2,
            "warning": 8, "error": 10, "critical": 15
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
        
        # Determina se aggiornare la statusBar
        update_condition = (
            elapsed_time >= self._last_status_min_time or
            current_level_idx > last_level_idx or
            level == self._last_status_level
        )
        
        if update_condition:
            # Per i vari livelli, prependi il tipo di messaggio se necessario
            prefix_map = {
                "error": "Errore: ", "critical": "Errore critico: ",
                "warning": "Attenzione: ", "loading": "Caricamento: "
            }
            prefix = prefix_map.get(level, "")
            status_message = f"{prefix}{message}"
            
            # Chiama il callback della barra di stato
            self.status_bar_callback(status_message)
            
            # Aggiorna le informazioni sull'ultimo stato
            self._last_status_level = level
            self._last_status_time = current_time
            self._last_status_min_time = min_display_seconds