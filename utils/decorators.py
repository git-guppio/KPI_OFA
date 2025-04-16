import logging
import functools
import traceback
import time
from typing import Callable, Optional, Any, Tuple, Type, Union

def error_logger(
    logger: Optional[logging.Logger] = None,
    log_level: int = logging.ERROR,
    include_traceback: bool = True,
    log_success: bool = True,
    success_level: int = logging.INFO,
    log_execution_time: bool = True
):
    """
    Decoratore avanzato per gestire le eccezioni e loggare informazioni.
    
    Args:
        logger: Logger da utilizzare. Se None, viene usato il logger root.
        log_level: Livello di logging per gli errori.
        include_traceback: Se includere il traceback degli errori nei log.
        log_success: Se loggare anche le esecuzioni di successo.
        success_level: Livello di logging per le esecuzioni di successo.
        log_execution_time: Se loggare il tempo di esecuzione della funzione.
    
    Returns:
        Un decoratore configurato secondo i parametri specificati.
    """
    # Se non viene fornito un logger, utilizziamo il logger root
    if logger is None:
        logger = logging.getLogger()
    
    def decorator(func: Callable) -> Callable:
        @functools.wraps(func)
        def wrapper(*args, **kwargs) -> Tuple[Any, Optional[Exception]]:
            func_name = func.__qualname__
            # Log dell'inizio dell'esecuzione
            if log_success:
                logger.log(success_level, f"Iniziata esecuzione di {func_name}")
            
            # Registra il tempo di inizio se richiesto
            start_time = time.time() if log_execution_time else None
            
            try:
                # Esecuzione della funzione
                result = func(*args, **kwargs)
                
                # Log del successo
                if log_success:
                    # Aggiungi info sul tempo di esecuzione se richiesto
                    if log_execution_time:
                        execution_time = time.time() - start_time
                        logger.log(success_level, 
                                  f"Completata esecuzione di {func_name} in {execution_time:.4f} secondi")
                    else:
                        logger.log(success_level, f"Completata esecuzione di {func_name}")
                
                return result, None
                
            except Exception as e:
                # Determinare il tipo di errore
                error_type = type(e).__name__
                error_msg = str(e)
                
                # Messaggio di log base
                log_message = f"Errore {error_type} in {func_name}: {error_msg}"
                
                # Aggiungi traceback se richiesto
                if include_traceback:
                    tb = traceback.format_exc()
                    log_message = f"{log_message}\n{tb}"
                
                # Log dell'errore
                logger.log(log_level, log_message)
                
                # Gestione specifica per alcuni tipi di errore
                if isinstance(e, ValueError):
                    logger.log(log_level - 10 if log_level > 10 else log_level, 
                              f"Errore di validazione: {error_msg}")
                elif isinstance(e, TypeError):
                    logger.log(log_level - 10 if log_level > 10 else log_level, 
                              f"Errore di tipo: {error_msg}")
                elif isinstance(e, IOError):
                    logger.log(log_level - 10 if log_level > 10 else log_level, 
                              f"Errore I/O: {error_msg}")
                
                return None, e
                
        return wrapper
    
    return decorator