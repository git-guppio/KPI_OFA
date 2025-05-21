#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Script di test rapido per la GUI dell'applicazione KPI OFA.

Questo script avvia l'interfaccia grafica dell'applicazione permettendo
di testare manualmente il suo funzionamento senza dover implementare
l'integrazione con SAP o altre dipendenze esterne.
"""

import os
import sys

# Aggiungi la directory principale al PYTHONPATH
current_dir = os.path.dirname(os.path.abspath(__file__))  # tests directory
project_root = os.path.dirname(current_dir)  # directory principale
sys.path.insert(0, project_root)


import json
import tempfile
import pandas as pd
import logging
#from kpi_ofa.core import setup_logging
from kpi_ofa.core.log_manager import LogManager
from kpi_ofa.ui.main_window import MainWindow


def setup_logging():
    """Configura il sistema di logging per l'applicazione"""
    log_dir = "logs"
    if not os.path.exists(log_dir):
        os.makedirs(log_dir)
        
    log_file = os.path.join(log_dir, "app.log")
    
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler(log_file),
            logging.StreamHandler()
        ]
    )
    
    # Logger principale
    logger = logging.getLogger("kpi_ofa")
    logger.setLevel(logging.DEBUG)
    
    return logger

# Configura il logging standard
setup_logging()

# Ottieni l'istanza del LogManager (singleton)
log_manager = LogManager()

# Registra un messaggio usando il LogManager
log_manager.log("Applicazione avviata da test_gui_manual.py", "info")



def setup_test_environment():
    """
    Configura l'ambiente di test creando file temporanei necessari.
    
    Returns:
        tuple: (config_file_path, excel_file_path, save_dir_path)
    """
    log_manager.log("Configurazione dell'ambiente di test...")
    
    # Crea una directory temporanea per i file di test
    temp_dir = tempfile.mkdtemp(prefix="kpi_ofa_test_")
    log_manager.log(f"Directory temporanea creata: {temp_dir}")
    
    # Crea una directory per il salvataggio dei dati
    save_dir = os.path.join(temp_dir, "save")
    os.makedirs(save_dir, exist_ok=True)
    
    # Crea un file di configurazione di test
    config_file = os.path.join(temp_dir, "config.json")
    config = {
        "save_directory": save_dir,
        "technologies": {
            "BESS": ["ITE", "USE", "CLE"],
            "SOLAR": ["ITS", "USS", "CLS", "BRS", "COS"],
            "WIND": ["ITW", "USW", "CLW", "BRW"]
        },
        "operations": {
            "estrai_AdM": True,
            "estrai_OdM": True,
            "estrai_AFKO": True,
            "elabora_xls": False
        }
    }
    
    with open(config_file, 'w', encoding='utf-8') as f:
        json.dump(config, f, indent=4)
    
    log_manager.log(f"File di configurazione creato: {config_file}")
    
    # Crea un file Excel di test con dati di esempio
    excel_file = os.path.join(temp_dir, "test_data.xlsx")
    
    # Crea un DataFrame con i dati di test
    data = {
        "idItem": [
            1000001, 2000001, 1000002, 2000002,
            "1000003-1", "2000003/1", "1000004", "2000004"
        ]
    }
    df = pd.DataFrame(data)
    
    # Salva il DataFrame in un file Excel
    with pd.ExcelWriter(excel_file, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name="AdM_OdM", index=False)
    
    log_manager.log(f"File Excel di test creato: {excel_file}")
    
    return config_file, excel_file, save_dir

def mock_sap_connection(*args, **kwargs):
    """
    Mock per la connessione SAP che simula sempre un successo.
    
    Returns:
        bool: True, simulando una connessione riuscita.
    """
    log_manager.log("Mock SAP: Simulazione connessione riuscita")
    return True

def mock_sap_extraction(*args, **kwargs):
    """
    Mock per l'estrazione SAP che simula sempre un successo.
    
    Returns:
        tuple: (True, DataFrame), simulando un'estrazione riuscita.
    """
    log_manager.log("Mock SAP: Simulazione estrazione dati riuscita")
    # Crea un DataFrame vuoto con le colonne attese
    df = pd.DataFrame(columns=["Ordine", "Descrizione", "Data", "Stato"])
    return True, df

def run_test_gui():
    """
    Avvia l'interfaccia grafica per il test manuale.
    """

    try:
        # Configura l'ambiente di test
        config_file, excel_file, save_dir = setup_test_environment()
        
        # Imposta le variabili d'ambiente per i test
        os.environ["KPI_OFA_CONFIG_FILE"] = config_file
        os.environ["KPI_OFA_TEST_EXCEL"] = excel_file
        
        # Importa solo dopo aver configurato l'ambiente
        from PyQt5.QtWidgets import QApplication
        
        # Aggiungi la directory principale del progetto al path
        sys.path.insert(0, os.path.abspath(os.path.dirname(os.path.dirname(__file__))))
        
        # Importa i moduli necessari
        from kpi_ofa.core.config_manager import ConfigManager
        from kpi_ofa.ui.main_window import MainWindow
        import kpi_ofa.constants as constants
        
        # Patch di constants per il test
        constants.configuration_json = config_file
        constants.required_sheet = "OnFieldApp.trace_kpi_user_actio"
        
        # Crea l'applicazione Qt
        app = QApplication(sys.argv)
        app.setApplicationName("KPI OFA Test")
        
        # Inizializza il gestore della configurazione
        config_manager = ConfigManager(config_file)
        
        # Crea la finestra principale
        window = MainWindow(config_manager)
        
        # Imposta il percorso del file Excel per test rapidi
        window.excel_file_path = excel_file
        window.file_text.setText(os.path.basename(excel_file))
        
        # Patch delle funzioni SAP per i test
        from kpi_ofa.services.sap_connection import SAPGuiConnection
        SAPGuiConnection.is_connected = mock_sap_connection
        
        from kpi_ofa.services.sap_transactions import SAPDataExtractor
        SAPDataExtractor.extract_IW29 = mock_sap_extraction
        SAPDataExtractor.extract_IW39 = mock_sap_extraction
        
        # Mostra la finestra
        window.show()
        
        print("\n" + "="*80)
        print("Test GUI di KPI OFA avviato")
        print(f"File di configurazione: {config_file}")
        print(f"File Excel di test: {excel_file}")
        print(f"Directory di salvataggio: {save_dir}")
        print("\nFunzionalità da testare:")
        print("1. Selezione del file Excel (già preimpostato)")
        print("2. Modifica delle date")
        print("3. Finestra di configurazione (pulsante 'Configura')")
        print("4. Pulsante 'Avvia' per simulare l'estrazione dati")
        print("5. Verifica dei messaggi di log")
        print("6. Pulsante 'Reset' per reimpostare l'interfaccia")
        print("="*80 + "\n")
        
        # Avvia il loop degli eventi
        exit_code = app.exec_()
        
        # Pulizia (opzionale per test manuali)
        # import shutil
        # shutil.rmtree(os.path.dirname(config_file))
        
        return exit_code
        
    except Exception as e:
        log_manager.log(f"Errore durante l'esecuzione del test: {e}", "error")
        return 1

if __name__ == "__main__":
    sys.exit(run_test_gui())