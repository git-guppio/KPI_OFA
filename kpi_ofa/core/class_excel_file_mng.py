import pandas as pd
import os
from pathlib import Path
import logging

# Configurazione logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

class ExcelFileMng:
    """
    Classe contenente strumenti per la gestione di file Excel.
    
    Funzionalità principali:
    - Caricamento automatico di file Excel di test con mapping predefinito
    - Creazione e gestione centralizzata di dataframe pandas
    - Salvataggio avanzato di dataframe in formato Excel
    - Validazione di percorsi e controllo esistenza file
    - Sistema di logging integrato per monitoraggio operazioni
    
    La classe utilizza un sistema di mapping predefinito per convertire file Excel
    specifici in dataframe con nomi standardizzati, facilitando l'uso in pipeline
    di test e analisi dati.
    
    Metodi disponibili:
    - __init__(base_path): Inizializza la classe con percorso base e mapping file
    - validate_path(): Valida l'esistenza e validità del percorso specificato
    - check_files_exist(): Verifica quali file del mapping esistono nella cartella
    - load_excel_file(filename): Carica un singolo file Excel in un DataFrame
    - load_all_files(): Carica tutti i file Excel mappati e crea dizionario dataframes
    - save_excel_file_advanced(): Salva DataFrame in Excel con opzioni avanzate
    
    Mapping file predefinito:
    - df_AFKO_test.xlsx → df_AFKO
    - df_excel_norm_test.xlsx → df_excel_normalized
    - IW29_AdM_test.xlsx → df_IW29  
    - IW39_OdM_test.xlsx → df_IW39
    - plants.xlsx → df_plants
    
    Utilizza il LogManager centralizzato per la registrazione dei messaggi
    sia nel sistema di logging che nell'interfaccia utente.
    """
    
    def __init__(self, base_path: str):
        """
        Inizializza il loader con il percorso base
        
        Args:
            base_path (str): Percorso della cartella contenente i file Excel
        """
        self.base_path = Path(base_path)
        self.test_file_to_df_mapping = {
            "df_AFKO_test.xlsx": "df_AFKO",
            "df_excel_norm_test.xlsx": "df_excel_normalized", 
            "IW29_AdM_test.xlsx": "df_IW29",
            "IW39_OdM_test.xlsx": "df_IW39",
            "plants.xlsx": "df_plants"
        }
        self.dataframes_dict = {}
    
    def validate_path(self) -> bool:
        """
        Valida l'esistenza del percorso specificato
        
        Returns:
            bool: True se il percorso esiste, False altrimenti
        """
        if not self.base_path.exists():
            logger.error(f"Il percorso {self.base_path} non esiste")
            return False
        
        if not self.base_path.is_dir():
            logger.error(f"Il percorso {self.base_path} non è una directory")
            return False
        
        logger.info(f"Percorso validato: {self.base_path}")
        return True
    
    def check_files_exist(self) -> dict:
        """
        Controlla quali file esistono nella cartella
        
        Returns:
            dict: Dizionario con file_name: exists (bool)
        """
        file_status = {}
        
        for filename in self.test_file_to_df_mapping.keys():
            file_path = self.base_path / filename
            exists = file_path.exists()
            file_status[filename] = exists
            
            if exists:
                logger.info(f"File trovato: {filename}")
            else:
                logger.warning(f"File mancante: {filename}")
        
        return file_status
    
    def load_excel_file(self, file_path: str, sheet_name: str = 0) -> tuple[bool, pd.DataFrame | None]:
        """
        Carica un singolo file Excel
        
        Args:
            file_path (str): Percorso completo del file da caricare
            sheet_name (str): Nome del foglio da caricare (default: primo foglio)
            
        Returns:
            pd.DataFrame: DataFrame caricato o DataFrame vuoto in caso di errore
        """
        
        try:
            # Prova a caricare il file Excel
            df = pd.read_excel(
                file_path,
                na_values=['', ' ', 'NULL', 'N/A'],  # Valori da considerare NaN
                sheet_name = sheet_name,
                header=0  # Modifica se header non è sulla prima riga
            )
            logger.info(f"File {file_path} caricato con successo - Shape: {df.shape}")
            return True, df
            
        except FileNotFoundError:
            logger.error(f"File non trovato: {file_path}")
            return False
            
        except Exception as e:
            logger.error(f"Errore durante il caricamento di {file_path}: {str(e)}")
            return False
    
    def load_all_files(self) -> tuple[bool, dict[str, pd.DataFrame] | None]:
        """
        Carica tutti i file Excel e crea i dataframe corrispondenti
        
        Returns:
            dict: Dizionario con nome_df: DataFrame
        """
        if not self.validate_path():
            return {}
        
        file_status = self.check_files_exist()
        num_df = 0 # numero di df presenti nel dizionario
        df_ok = 0 # numero di df caricati correttamente
        for filename, df_name in self.test_file_to_df_mapping.items():
            num_df += 1
            if file_status.get(filename, False):
                result, df = self.load_excel_file(filename)
                if result:
                    df_ok += 1
                    self.dataframes_dict[df_name] = df
                    logger.info(f"DataFrame {df_name} creato con successo")
                else:
                    logger.warning(f"DataFrame {df_name} è vuoto")
            else:
                logger.warning(f"Saltato {filename} - file non trovato")
        
        if num_df == df_ok:
            logger.info(f"Tutti i {num_df} DataFrame sono stati caricati correttamente")
            return True, self.dataframes_dict    
        else:
            logger.error(f"Caricamento fallito: {df_ok}/{num_df} DataFrame caricati correttamente")
            return False, None

    def save_excel_file_advanced(self, df: pd.DataFrame, filename: str, 
                            sheet_name: str = 'Sheet1', 
                            index: bool = False,
                            overwrite: bool = True) -> bool:
        """
        Salva un DataFrame in un file Excel con opzioni avanzate
        
        Args:
            df (pd.DataFrame): DataFrame da salvare
            filename (str): Nome del file da creare/sovrascrivere
            sheet_name (str): Nome del foglio Excel (default: 'Sheet1')
            index (bool): Se includere l'indice come colonna (default: False)
            overwrite (bool): Se sovrascrivere file esistenti (default: True)
            
        Returns:
            bool: True se salvato con successo, False in caso di errore
        """
        # Verifico se la directory esiste e solo in tal caso salvo il file       
        file_path = self.base_path / filename
        if not self.validate_path():
            return False
        
        try:
            # Verifica che il DataFrame non sia vuoto
            if df.empty:
                logger.warning(f"DataFrame vuoto, salvataggio di {filename} non eseguiuto")
                return False
            
            # Controlla se il file esiste già
            if file_path.exists() and not overwrite:
                logger.warning(f"File {filename} già esistente, salvataggio non eseguito")
                return False
            
            # # Crea la directory se non esiste
            # file_path.parent.mkdir(parents=True, exist_ok=True)
            
            # Salva il DataFrame in Excel
            df.to_excel(
                file_path,
                sheet_name=sheet_name,
                index=index,
                na_rep='',
                header=True,
                engine='openpyxl'  # Engine specifico per .xlsx
            )
            
            logger.info(f"File {filename} salvato con successo - Shape: {df.shape} - Path: {file_path}")
            return True
            
        except PermissionError:
            logger.error(f"Permessi insufficienti per scrivere il file: {filename}")
            return False
            
        except FileNotFoundError:
            logger.error(f"Percorso non trovato: {file_path.parent}")
            return False
            
        except Exception as e:
            logger.error(f"Errore durante il salvataggio di {filename}: {str(e)}")
            return False