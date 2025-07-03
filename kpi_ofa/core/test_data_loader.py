# kpi_ofa/core/test_data_loader.py

import os
import pandas as pd
from pathlib import Path
from typing import Dict, Optional
from datetime import datetime

import kpi_ofa.constants as constants


class TestDataLoadError(Exception):
    """Eccezione personalizzata per errori di caricamento file di test."""
    pass


class TestDataLoader:
    """
    Classe per il caricamento dei file di test.
    
    Carica i file Excel di test e assegna i DataFrame agli attributi dell'istanza.
    Se anche un solo file fallisce, solleva un'eccezione.
    """
    
    def __init__(self, main_window=None):
        """
        Inizializza il TestDataLoader.
        
        Args:
            main_window: Riferimento all'istanza di MainWindow per callback UI
        """
        self.main_window = main_window
        self.config = None
        self.file_to_df_mapping = constants.test_file_to_df_mapping
        
        # Inizializza DataFrame vuoti
        self._init_empty_dataframes()
        
        self.log("TestDataLoader inizializzato", "info")
    
    def _init_empty_dataframes(self):
        """Inizializza DataFrame vuoti per tutti i file di test."""
        for filename, df_name in self.file_to_df_mapping.items():
            setattr(self, df_name, pd.DataFrame())
    
    def set_config(self, config: Dict):
        """
        Imposta la configurazione.
        
        Args:
            config: Dizionario di configurazione
        """
        self.config = config
        self.log("Configurazione aggiornata in TestDataLoader", "info")
    
    def load_test_files(self) -> bool:
        """
        Carica tutti i file di test.
        
        Carica tutti i file Excel dalla directory test/ e li assegna come attributi.
        Se anche un solo file fallisce, solleva TestDataLoadError.
        
        Returns:
            bool: True se tutti i file sono stati caricati con successo
            
        Raises:
            TestDataLoadError: Se il caricamento di uno o più file fallisce
        """
        if not constants.DEBUG_MODE:
            self.log("Debug mode OFF - caricamento file di test saltato", "info")
            return False
        
        self.log("🔧 Caricamento file di test", "info")
        
        # Verifica prerequisiti
        self._validate_prerequisites()
        
        # Carica tutti i file
        test_dir = self._get_test_directory()
        self._load_all_files(test_dir)
        
        # Se arriviamo qui, tutto è andato bene
        self.log("✅ Tutti i file di test caricati con successo", "success")
        self.update_start_button_state(True)
        return True
    
    def _validate_prerequisites(self):
        """
        Valida i prerequisiti per il caricamento.
        
        Raises:
            TestDataLoadError: Se i prerequisiti non sono soddisfatti
        """
        if not self.config:
            raise TestDataLoadError("Configurazione non disponibile")
        
        test_dir = constants.test_directory
        if not os.path.exists(test_dir):
            raise TestDataLoadError(f"Directory test non trovata: {test_dir}")
    
    def _get_test_directory(self) -> str:
        """Ottiene il percorso della directory test."""
        save_dir = constants.test_directory
        return save_dir
    
    def _load_all_files(self, test_dir: str):
        """
        Carica tutti i file dalla directory test.
        
        Args:
            test_dir: Percorso della directory test
            
        Raises:
            TestDataLoadError: Se il caricamento di un file fallisce
        """
        total_files = len(self.file_to_df_mapping)
        loaded_dataframes = {}
        
        for idx, (filename, df_name) in enumerate(self.file_to_df_mapping.items()):
            # Progress update
            progress = int((idx / total_files) * 100)
            self.log(f"📂 Caricamento {filename}... ({idx+1}/{total_files})", "loading", 
                    update_status=True, progress=progress)
            
            file_path = os.path.join(test_dir, filename)
            
            # Carica il file
            df = self._load_single_file(file_path, filename)
            loaded_dataframes[df_name] = df
            
            self.log(f"✅ {filename}: {len(df)} righe, {len(df.columns)} colonne", "success")
        
        # Se arriviamo qui, tutti i file sono stati caricati
        self._assign_dataframes(loaded_dataframes)
    
    def _load_single_file(self, file_path: str, filename: str) -> pd.DataFrame:
        """
        Carica un singolo file Excel.
        
        Args:
            file_path: Percorso completo del file
            filename: Nome del file (per logging)
            
        Returns:
            pd.DataFrame: DataFrame caricato
            
        Raises:
            TestDataLoadError: Se il caricamento fallisce
        """
        try:
            if not os.path.exists(file_path):
                raise TestDataLoadError(f"File non trovato: {filename}")
            
            # Carica il file Excel
            df = pd.read_excel(
                file_path,
                engine='openpyxl' if file_path.endswith('.xlsx') else 'xlrd'
            )
            
            # Verifica che non sia vuoto (opzionale - rimuovi se vuoi accettare file vuoti)
            if df.empty:
                self.log(f"⚠️  File {filename} è vuoto", "warning")
            
            return df
            
        except FileNotFoundError:
            raise TestDataLoadError(f"File non trovato: {filename}")
        except PermissionError:
            raise TestDataLoadError(f"Permessi insufficienti per leggere: {filename}")
        except pd.errors.EmptyDataError:
            raise TestDataLoadError(f"File Excel vuoto o corrotto: {filename}")
        except Exception as e:
            raise TestDataLoadError(f"Errore caricamento {filename}: {str(e)}")
    
    def _assign_dataframes(self, dataframes: Dict[str, pd.DataFrame]):
        """
        Assegna i DataFrame agli attributi dell'istanza.
        
        Args:
            dataframes: Dizionario con i DataFrame da assegnare
        """
        for df_name, df in dataframes.items():
            setattr(self, df_name, df)
            self.log(f"📋 Assegnato {df_name}: {len(df)} righe", "info")
    
    def get_dataframe(self, df_name: str) -> Optional[pd.DataFrame]:
        """
        Ottiene un DataFrame specifico.
        
        Args:
            df_name: Nome del DataFrame da ottenere
            
        Returns:
            DataFrame richiesto o None se non trovato
        """
        return getattr(self, df_name, None)
    
    def is_loaded(self) -> bool:
        """
        Verifica se TUTTI i DataFrame sono stati caricati con dati validi.
        
        Returns:
            True se TUTTI i DataFrame esistono e hanno almeno una riga
        """
        for df_name in self.file_to_df_mapping.values():
            df = getattr(self, df_name, pd.DataFrame())
            if df.empty or len(df) == 0:
                return False
        return True
    
    def clear_data(self):
        """Pulisce tutti i DataFrame caricati."""
        for df_name in self.file_to_df_mapping.values():
            setattr(self, df_name, pd.DataFrame())
        self.log("🗑️  Dati test puliti", "info")
    
    def log(self, message: str, level: str = "info", update_status: bool = False, 
            update_log: bool = True, progress: int = 0):
        """
        Delega il logging alla MainWindow se disponibile.
        
        Args:
            message: Messaggio da loggare
            level: Livello del log
            update_status: Se aggiornare la status bar
            update_log: Se aggiornare il widget di log
            progress: Valore di progresso (0-100)
        """
        if self.main_window and hasattr(self.main_window, 'log_manager'):
            self.main_window.log_manager.log(
                message, level, 
                update_status=update_status, 
                update_log=update_log
            )
        else:
            print(f"[{level.upper()}] {message}")
    
    def update_start_button_state(self, enabled: bool):
        """
        Aggiorna lo stato del pulsante Start tramite MainWindow.
        
        Args:
            enabled: True per abilitare il pulsante, False per disabilitarlo
        """
        if self.main_window and hasattr(self.main_window, 'update_start_button_state'):
            self.main_window.update_start_button_state(enabled)