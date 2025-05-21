#!/usr/bin/env python3
"""
Entry point principale per l'applicazione kpi_ofa.
"""
import sys
import os
import logging
from PyQt5.QtWidgets import QApplication

def setup_logging():
    """Configura il sistema di logging per l'applicazione"""
    # Crea la directory logs se non esiste
    log_dir = "logs"
    if not os.path.exists(log_dir):
        os.makedirs(log_dir)
        
    # Percorso del file di log
    log_file = os.path.join(log_dir, "app.log")
    
    # Configura il logging di base
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler(log_file),
            logging.StreamHandler()
        ]
    )
    
    # Logger principale dell'applicazione
    logger = logging.getLogger("kpi_ofa")
    logger.setLevel(logging.DEBUG)
    
    return logger

def main():
    """Funzione principale dell'applicazione"""
    # Configura il logging
    logger = setup_logging()
    logger.info("Applicazione avviata da main.py")
    
    # Importa i moduli necessari
    # Importiamo qui per evitare problemi di circolarità delle dipendenze
    from kpi_ofa.core.config_manager import ConfigManager
    from kpi_ofa.ui.main_window import MainWindow
    import kpi_ofa.constants as constants
    
    # Crea l'applicazione Qt
    app = QApplication(sys.argv)
    app.setApplicationName("KPI OFA")
    
    # Inizializza il gestore della configurazione
    config_manager = ConfigManager(constants.configuration_json)
    
    # Crea la finestra principale
    window = MainWindow(config_manager)
    window.show()  # Mostra la finestra principale
    
    # Avvia il loop degli eventi di Qt
    return app.exec_()

if __name__ == "__main__":
    sys.exit(main())