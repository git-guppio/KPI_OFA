# kpi_ofa/core/excel_data_processor.py

import os
import pandas as pd
import logging
from typing import Tuple, List, Dict, Optional, Union

from kpi_ofa.core.log_manager import LogManager

logger = logging.getLogger(__name__)

class ExcelDataProcessor:
    """
    Processore dei dati per l'applicazione KPI OFA.
    
    Questa classe è responsabile dell'elaborazione dei dati dal file Excel
    e della preparazione degli Avvisi di Manutenzione (AdM) e Ordini di 
    Manutenzione (OdM) per l'estrazione SAP.
    
    Utilizza il LogManager centralizzato per la registrazione dei messaggi
    sia nel sistema di logging che nell'interfaccia utente.
    """
    
    def __init__(self):
        """Inizializza il processore di dati."""
        # DataFrame contenenti i dati elaborati
        self.df_AdM = None
        self.df_OdM = None
        self.df_normalized = None
        
        # Percorso del file Excel selezionato
        self.excel_file_path = None
        
        # Ottieni l'istanza del LogManager (singleton)
        self.log_manager = LogManager()
        
        self.log("ExcelDataProcessor inizializzato")
    
    def log(self, message: str, level: str = "info", update_status: bool = True, update_log: bool = True):
        """
        Registra un messaggio tramite il LogManager centralizzato.
        
        Args:
            message (str): Messaggio da registrare.
            level (str): Livello del log ('info', 'warning', 'error', 'success', 'loading').
            update_status (bool): Se aggiornare la barra di stato.
            update_log (bool): Se aggiornare il widget di log.
        """
        # Utilizza il LogManager per registrare il messaggio
        self.log_manager.log(message, level, update_status, update_log, origin="excel_data_processor")
    
    def process_excel_file(self, file_path: str, required_sheet: str, required_columns: List[str]) -> bool:
        """
        Processa il file Excel, verificandone la struttura e estraendo AdM e OdM.
        
        Args:
            file_path (str): Percorso del file Excel.
            required_sheet (str): Nome dello sheet richiesto.
            required_columns (list): Lista delle colonne richieste.
            
        Returns:
            bool: True se il processing è riuscito, False altrimenti.
        """
        self.log(f"Elaborazione del file Excel: {file_path}")
        
        # Salva il percorso del file
        self.excel_file_path = file_path
        
        # Verifica del file Excel
        self.log("Verifica struttura file Excel", "loading")
        result, df = self.check_excel_file(file_path, required_sheet, required_columns)
        
        if not result:
            self.log("Verifica struttura file Excel fallita", "error")
            return False
            
        self.log("Struttura file Excel verificata con successo", "success")
        
        # Normalizzazione dei dati
        self.log("Normalizzazione degli idItem", "loading")
        result, self.df_normalized = self.normalize_df(df)
        if not result:
            self.log("Normalizzazione degli idItem fallita", "error")
            return False
        
        # Estrazione AdM
        self.log("Estrazione degli Avvisi di Manutenzione (AdM)", "loading")
        result, self.df_AdM = self.extract_adm(self.df_normalized)
        if not result:
            self.log("Estrazione degli AdM fallita", "error")
            return False
            
        self.log(f"AdM estratti: {len(self.df_AdM)}", "success")
        
        # Estrazione OdM
        self.log("Estrazione degli Ordini di Manutenzione (OdM)", "loading")
        result, self.df_OdM = self.extract_odm(self.df_normalized)
        if not result:
            self.log("Estrazione degli OdM fallita", "error")
            return False
            
        self.log(f"OdM estratti: {len(self.df_OdM)}", "success")
        self.log("File Excel elaborato con successo", "success")
        
        return True
    
    def check_excel_file(self, file_path: str, required_sheet: str, 
                         required_columns: List[str]) -> Tuple[bool, Optional[pd.DataFrame]]:
        """
        Verifica se il file Excel contiene lo sheet e le colonne richieste.
        
        Args:
            file_path (str): Percorso del file Excel.
            required_sheet (str): Nome dello sheet richiesto.
            required_columns (list): Lista delle colonne richieste.
            
        Returns:
            tuple: (bool, DataFrame) - True se il file è valido, il DataFrame se disponibile.
        """
        try:
            self.log(f"Verifica del file Excel: {file_path}")
            
            # Controllo esistenza file
            if not os.path.exists(file_path):
                self.log(f"Il file non esiste: {file_path}", "error")
                return False, None
            
            # Ottieni la lista degli sheet nel file
            excel_file = pd.ExcelFile(file_path)
            sheet_names = excel_file.sheet_names
            self.log(f"Sheet presenti nel file: {', '.join(sheet_names)}")
            
            # Verifica se lo sheet esiste
            if required_sheet in sheet_names:
                self.log(f"Sheet '{required_sheet}' trovato nel file", "success")
                # Leggi i dati dallo sheet
                df = pd.read_excel(file_path, sheet_name=required_sheet)
                
                self.log(f"Colonne presenti nel file: {', '.join(df.columns)}")
                
                # Verifica se tutte le colonne richieste sono presenti
                missing_columns = [col for col in required_columns if col not in df.columns]
                
                if missing_columns:
                    # Alcune colonne sono mancanti
                    missing_cols_str = ", ".join(missing_columns)
                    self.log(f"Colonne mancanti: {missing_cols_str}", "error")
                    self.log(f"Colonne disponibili: {', '.join(df.columns)}", "error")
                    return False, None
                else:
                    # Tutte le colonne sono presenti
                    rows_count = len(df)
                    self.log(f"File valido: {rows_count} righe trovate", "success")
                    return True, df
            else:
                # Lo sheet non è presente
                self.log(f"Lo sheet '{required_sheet}' non è presente nel file", "error")
                return False, None
                
        except pd.errors.EmptyDataError:
            self.log("Il file Excel è vuoto", "error")
            return False, None
        except pd.errors.ParserError:
            self.log("Errore nella lettura del file Excel: formato non valido", "error")
            return False, None
        except Exception as e:
            self.log(f"Errore nella lettura del file Excel: {str(e)}", "error")
            return False, None
    
    def normalize_df(self, df: pd.DataFrame) -> Tuple[bool, Optional[pd.DataFrame]]:
        """
        Normalizza il DataFrame, elaborando i dati nella colonna 'idItem'.
        
        Estrae il valore numerico base dagli idItem, considerando formati come:
        - Numeri semplici (es. 1000001)
        - Stringhe con trattino (es. "1000003-1")
        - Stringhe con slash (es. "2000003/1")
        
        Args:
            df (pd.DataFrame): DataFrame originale.
            
        Returns:
            tuple: (bool, DataFrame) - True se la normalizzazione è riuscita, DataFrame normalizzato.
        """
        try:
            self.log("Normalizzazione della colonna idItem del DataFrame")
            self.log(f"Presenti {len(df)} idItem nel DataFrame")
            
            # Crea un nuovo DataFrame con solo la colonna idItem
            id_items_df = df[["idItem"]].copy()
            
            # Verifica che il DataFrame non sia vuoto
            if id_items_df.empty:
                self.log("Nessun elemento trovato nel file Excel", "error")
                return False, None
            
            # Rimuovi eventuali valori nulli
            id_items_df_notnull = id_items_df.dropna(subset=["idItem"])
            if len(id_items_df_notnull) < len(id_items_df):
                self.log(f"Rimossi {len(id_items_df) - len(id_items_df_notnull)} valori nulli", "warning")
            
            id_items_df = id_items_df_notnull
            
            # Funzione per estrarre la parte prima del primo - o / e convertire in intero
            def extract_base_id(id_text):
                if not isinstance(id_text, str):
                    try:
                        # Se è già un numero, prova a convertirlo direttamente
                        return int(id_text)
                    except (ValueError, TypeError):
                        logger.warning(f"Impossibile convertire '{id_text}' in intero")
                        return None
                        
                # Cerca il primo trattino o slash
                dash_pos = id_text.find('-')
                slash_pos = id_text.find('/')
                
                # Determina quale carattere appare per primo (se presente)
                if dash_pos >= 0 and (slash_pos < 0 or dash_pos < slash_pos):
                    base_id = id_text[:dash_pos]
                elif slash_pos >= 0:
                    base_id = id_text[:slash_pos]
                else:
                    base_id = id_text  # Nessun trattino o slash trovato
                    
                # Converti in intero se possibile
                try:
                    return int(base_id)
                except ValueError:
                    logger.warning(f"Impossibile convertire '{base_id}' in intero")
                    return None
            
            # Crea una nuova colonna con i valori normalizzati
            id_items_df['idItem_base'] = id_items_df['idItem'].apply(extract_base_id)
            
            # Rimuovi righe con valori None nella colonna idItem_base
            id_items_df_clean = id_items_df.dropna(subset=['idItem_base'])
            if len(id_items_df_clean) < len(id_items_df):
                self.log(f"Rimossi {len(id_items_df) - len(id_items_df_clean)} valori non convertibili", "warning")
            
            id_items_df = id_items_df_clean
            
            # Usa la colonna normalizzata come colonna principale
            id_items_df['idItem'] = id_items_df['idItem_base']
            id_items_df = id_items_df[['idItem']]  # Mantieni solo la colonna idItem
            
            # Rimuovi eventuali duplicati
            id_items_df_unique = id_items_df.drop_duplicates()
            if len(id_items_df_unique) < len(id_items_df):
                self.log(f"Rimossi {len(id_items_df) - len(id_items_df_unique)} valori duplicati", "info")
            
            id_items_df = id_items_df_unique
            
            # Resetta l'indice
            id_items_df = id_items_df.reset_index(drop=True)
            
            self.log(f"Estratti {len(id_items_df)} idItem unici come valori interi", "success")
            if (len(id_items_df) != len(df)):
                self.log(f"{len(df) - len(id_items_df)} righe eliminate durante la normalizzazione", "info")
                
            return True, id_items_df
            
        except Exception as e:
            self.log(f"Errore nell'estrazione degli idItem: {str(e)}", "error")
            logger.error(f"Dettaglio errore: {str(e)}", exc_info=True)
            return False, None
    
    def extract_adm(self, df: pd.DataFrame) -> Tuple[bool, Optional[pd.DataFrame]]:
        """
        Estrae gli Avvisi di Manutenzione (AdM) dal DataFrame normalizzato.
        
        Gli AdM sono identificati da idItem < 2000000000.
        
        Args:
            df (pd.DataFrame): DataFrame normalizzato.
            
        Returns:
            tuple: (bool, DataFrame) - True se l'estrazione è riuscita, DataFrame degli AdM.
        """
        try:
            self.log(f"Estrazione AdM - Presenti {len(df)} idItem nel DataFrame")
            
            # Assicurati che i valori siano numerici
            df_numeric = df.copy()
            df_numeric["idItem"] = pd.to_numeric(df_numeric["idItem"], errors="coerce")
            
            # Rimuovi le righe con valori NaN
            df_numeric_clean = df_numeric.dropna(subset=["idItem"])
            
            if len(df_numeric_clean) < len(df_numeric):
                self.log(f"Rimossi {len(df_numeric) - len(df_numeric_clean)} valori non numerici", "warning")
            
            df_numeric = df_numeric_clean
            
            self.log(f"Presenti {len(df_numeric)} idItem numerici nel DataFrame")
            
            # Filtra il DataFrame per ottenere solo i valori minori di 2000000000
            AdM_df = df_numeric[df_numeric["idItem"] < 2000000000]
            
            # Resetta l'indice
            AdM_df = AdM_df.reset_index(drop=True)
            
            self.log(f"Filtrati {len(AdM_df)} record con idItem < 2000000000 su {len(df_numeric)} totali", "success")
            
            return True, AdM_df
            
        except Exception as e:
            self.log(f"Errore nell'estrazione degli AdM: {str(e)}", "error")
            logger.error(f"Dettaglio errore: {str(e)}", exc_info=True)
            return False, None
    
    def extract_odm(self, df: pd.DataFrame) -> Tuple[bool, Optional[pd.DataFrame]]:
        """
        Estrae gli Ordini di Manutenzione (OdM) dal DataFrame normalizzato.
        
        Gli OdM sono identificati da idItem >= 2000000000.
        
        Args:
            df (pd.DataFrame): DataFrame normalizzato.
            
        Returns:
            tuple: (bool, DataFrame) - True se l'estrazione è riuscita, DataFrame degli OdM.
        """
        try:
            self.log(f"Estrazione OdM - Presenti {len(df)} idItem nel DataFrame")
            
            # Assicurati che i valori siano numerici
            df_numeric = df.copy()
            df_numeric["idItem"] = pd.to_numeric(df_numeric["idItem"], errors="coerce")
            
            # Rimuovi le righe con valori NaN
            df_numeric_clean = df_numeric.dropna(subset=["idItem"])
            
            if len(df_numeric_clean) < len(df_numeric):
                self.log(f"Rimossi {len(df_numeric) - len(df_numeric_clean)} valori non numerici", "warning")
            
            df_numeric = df_numeric_clean
            
            self.log(f"Presenti {len(df_numeric)} idItem numerici nel DataFrame")
            
            # Filtra il DataFrame per ottenere solo i valori maggiori o uguali a 2000000000
            OdM_df = df_numeric[df_numeric["idItem"] >= 2000000000]
            
            # Resetta l'indice
            OdM_df = OdM_df.reset_index(drop=True)
            
            self.log(f"Filtrati {len(OdM_df)} record con idItem >= 2000000000 su {len(df_numeric)} totali", "success")
            
            return True, OdM_df
            
        except Exception as e:
            self.log(f"Errore nell'estrazione degli OdM: {str(e)}", "error")
            logger.error(f"Dettaglio errore: {str(e)}", exc_info=True)
            return False, None
    
    def get_adm(self) -> Optional[pd.DataFrame]:
        """
        Restituisce il DataFrame degli Avvisi di Manutenzione (AdM).
        
        Returns:
            pd.DataFrame: DataFrame degli AdM o None se non disponibile.
        """
        return self.df_AdM
    
    def get_odm(self) -> Optional[pd.DataFrame]:
        """
        Restituisce il DataFrame degli Ordini di Manutenzione (OdM).
        
        Returns:
            pd.DataFrame: DataFrame degli OdM o None se non disponibile.
        """
        return self.df_OdM
    
    def save_processed_data(self, save_dir: str) -> bool:
        """
        Salva i dati elaborati in file Excel separati.
        
        Args:
            save_dir (str): Directory in cui salvare i file.
            
        Returns:
            bool: True se il salvataggio è riuscito, False altrimenti.
        """
        try:
            # Verifica che la directory esista
            if not os.path.exists(save_dir):
                os.makedirs(save_dir, exist_ok=True)
                self.log(f"Creata directory: {save_dir}", "info")
            
            # Verifica che ci siano dati da salvare
            if self.df_AdM is None or self.df_OdM is None:
                self.log("Nessun dato da salvare. Elaborare prima un file Excel.", "error")
                return False
            
            # Salva gli AdM
            adm_file = os.path.join(save_dir, "AdM_estratti.xlsx")
            self.df_AdM.to_excel(adm_file, index=False)
            self.log(f"File AdM salvato: {adm_file}", "success")
            
            # Salva gli OdM
            odm_file = os.path.join(save_dir, "OdM_estratti.xlsx")
            self.df_OdM.to_excel(odm_file, index=False)
            self.log(f"File OdM salvato: {odm_file}", "success")
            
            return True
        except Exception as e:
            self.log(f"Errore nel salvataggio dei dati elaborati: {str(e)}", "error")
            logger.error(f"Dettaglio errore: {str(e)}", exc_info=True)
            return False
    
    def get_statistics(self) -> Dict[str, Union[int, float, str]]:
        """
        Restituisce statistiche sui dati elaborati.
        
        Returns:
            dict: Dizionario con statistiche sui dati.
        """
        stats = {
            "file_elaborato": os.path.basename(self.excel_file_path) if self.excel_file_path else "Nessuno",
            "totale_righe_originali": len(self.df_normalized) if self.df_normalized is not None else 0,
            "num_adm": len(self.df_AdM) if self.df_AdM is not None else 0,
            "num_odm": len(self.df_OdM) if self.df_OdM is not None else 0
        }
        
        # Calcola statistiche percentuali se ci sono dati
        if stats["totale_righe_originali"] > 0:
            stats["percentuale_adm"] = round(stats["num_adm"] / stats["totale_righe_originali"] * 100, 2)
            stats["percentuale_odm"] = round(stats["num_odm"] / stats["totale_righe_originali"] * 100, 2)
        else:
            stats["percentuale_adm"] = 0
            stats["percentuale_odm"] = 0
        
        return stats
    
    def clear_data(self) -> None:
        """Pulisce tutti i dati elaborati."""
        self.df_AdM = None
        self.df_OdM = None
        self.df_normalized = None
        self.excel_file_path = None
        
        self.log("Dati elaborati cancellati", "info")