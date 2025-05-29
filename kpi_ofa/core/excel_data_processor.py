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
        self.df_excel = pd.DataFrame()
        self.df_AdM = pd.DataFrame()
        self.df_OdM = pd.DataFrame()
        self.df_excel_normalized = pd.DataFrame()
        
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
    
    def process_excel_file(self, file_path: str, required_sheet: str, required_columns: List[str]) -> Tuple[bool, Optional[pd.DataFrame]]:
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
        # salvo il df per poterlo usare in altre funzioni
        self.df_excel = df
        
        # Normalizzazione dei dati -> Creo la colonna idItem_normalized contenente i valori della colonna idItem 
        self.log("Normalizzazione degli idItem", "loading")
        result, self.df_excel_normalized = self.normalize_excel_df(df)
        if not result:
            self.log("Normalizzazione degli idItem fallita", "error")
            return False
        
        # Aggiungo la colonna AdM e OdM al DataFrame normalizzato
        self.df_excel_normalized["AdM"] = self.df_excel_normalized["idItem_normalized"].where(self.df_excel_normalized["idItem_normalized"] < 2000000000)
        
        self.df_excel_normalized["OdM"] = self.df_excel_normalized["idItem_normalized"].where(self.df_excel_normalized["idItem_normalized"] > 2000000000)
        
        # Estrazione AdM
        self.log("Estrazione degli Avvisi di Manutenzione (AdM)", "loading")
        result, self.df_AdM = self.extract_adm(self.df_excel_normalized)
        if not result:
            self.log("Estrazione degli AdM fallita", "error")
            return False
            
        self.log(f"AdM estratti: {len(self.df_AdM)}", "success")
        
        # Estrazione OdM
        self.log("Estrazione degli Ordini di Manutenzione (OdM)", "loading")
        result, self.df_OdM = self.extract_odm(self.df_excel_normalized)
        if not result:
            self.log("Estrazione degli OdM fallita", "error")
            return False
        
        self.log(f"OdM estratti: {len(self.df_OdM)}", "success")

        # Creo le colonne 'country' e 'tecnologia' nel dataframe df_excel_normalized
        # Crea colonna TL (terzo carattere)
        def extract_and_transform_tech(functional_location):
            """Estrae il primo carattere e lo trasforma in tecnologia"""
            if pd.isna(functional_location) or not isinstance(functional_location, str) or len(functional_location) == 0:
                return None
            
            first_char = functional_location[0].upper()
            
            mapping = {
                'S': 'Solar',
                'E': 'Bess',
                'W': 'WIND'
            }
            
            return mapping.get(first_char, None)

        # Applica tutto insieme
        self.df_excel_normalized['tecnologia'] = self.df_excel_normalized['functionalLocation'].apply(extract_and_transform_tech)

        # Crea colonna tecnologia (primi due caratteri)
        self.df_excel_normalized['country'] = self.df_excel_normalized['functionalLocation'].apply(
            lambda x: x[:2] if pd.notna(x) and isinstance(x, str) and len(x) >= 2 else None
        )

        self.log("File Excel elaborato con successo", "success")

        return True, self.df_excel_normalized
    
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
    
    def normalize_excel_df(self, df: pd.DataFrame) -> Tuple[bool, Optional[pd.DataFrame]]:
        """
        Normalizza il DataFrame, elaborando i dati nella colonna 'idItem'.
        
        Crea una nuova colonna 'idItem_normalized' con i valori base estratti,
        mantenendo tutte le righe originali (anche quelle con valori vuoti).
        
        Estrae il valore numerico base dagli idItem, considerando formati come:
        - Numeri semplici (es. 1000001)
        - Stringhe con trattino (es. "1000003-1" → 1000003)
        - Stringhe con slash (es. "2000003/1" → 2000003)
        - Formati complessi (es. "240000493108-0010-2001113622" → 240000493108)
        
        Args:
            df (pd.DataFrame): DataFrame originale.
            
        Returns:
            tuple: (bool, DataFrame) - True se la normalizzazione è riuscita, DataFrame con nuova colonna.
        """
        try:
            self.log("Normalizzazione della colonna idItem del DataFrame")
            self.log(f"Presenti {len(df)} righe nel DataFrame")
            
            # Lavora su una copia del DataFrame originale
            df_result = df.copy()
            
            # Verifica che il DataFrame non sia vuoto
            if df_result.empty:
                self.log("DataFrame vuoto", "error")
                return False, None
            
            # Verifica che la colonna idItem esista
            if 'idItem' not in df_result.columns:
                self.log("Colonna 'idItem' non trovata nel DataFrame", "error")
                return False, None
            
            def is_valid_iditem(value):
                """Controlla se un idItem è valido per la normalizzazione"""
                if pd.isna(value):
                    return False
                str_val = str(value).strip()
                return str_val not in ['', 'nan', 'None', '0']
           
            # Rimuove eventuali righe dal DF in cui idItem è NaN o vuoto oppure vale 0
            self.log("Rimozione righe con idItem vuoto o NaN", "info")
            num_righe = len(df_result)
            df_clean = df_result[df_result['idItem'].apply(is_valid_iditem)].copy()
            if num_righe != len(df_clean):
                self.log(f"Rimosse {num_righe - len(df_clean)} righe non valide", "warning")
            self.log(f"Righe rimanenti dopo rimozione: {len(df_clean)}", "info")
            df_result = df_clean

            # Funzione per estrarre la parte prima del primo - o / e convertire in intero
            def extract_base_id(id_text):
                """
                Estrae il valore base da un idItem
                
                Esempi:
                - 240000493138/0010 → 240000493138
                - 240000493124-0010 → 240000493124  
                - 240000493108-0010-2001113622 → 240000493108
                - 1000001 → 1000001
                - None → None
                - "" → None
                """
                # Gestisci valori nulli o vuoti
                if pd.isna(id_text) or id_text == '' or id_text is None:
                    return None
                    
                # Se non è una stringa, prova a convertirlo
                if not isinstance(id_text, str):
                    try:
                        # Se è già un numero, prova a convertirlo direttamente
                        return int(id_text)
                    except (ValueError, TypeError):
                        self.log(f"Impossibile convertire '{id_text}' in intero", "warning")
                        return None
                
                # Converte in stringa per sicurezza
                id_str = str(id_text).strip()
                
                # Se stringa vuota dopo strip
                if not id_str:
                    return None
                    
                # Cerca il primo trattino o slash
                dash_pos = id_str.find('-')
                slash_pos = id_str.find('/')
                
                # Determina quale carattere appare per primo (se presente)
                if dash_pos >= 0 and (slash_pos < 0 or dash_pos < slash_pos):
                    base_id = id_str[:dash_pos]
                elif slash_pos >= 0:
                    base_id = id_str[:slash_pos]
                else:
                    base_id = id_str  # Nessun trattino o slash trovato
                    
                # Converti in intero se possibile
                try:
                    return int(base_id)
                except ValueError:
                    self.log(f"Impossibile convertire '{base_id}' (da '{id_text}') in intero", "warning")
                    return None

            # Applica la normalizzazione e crea la nuova colonna
            self.log("Applicazione della normalizzazione...")
            df_result['idItem_normalized'] = df_result['idItem'].apply(extract_base_id)

            # Statistiche sulla normalizzazione
            righe_totali = len(df_result)
            valori_originali_nulli = df_result['idItem'].isna().sum()
            valori_originali_vuoti = (df_result['idItem'] == '').sum()
            valori_normalizzati_nulli = df_result['idItem_normalized'].isna().sum()
            valori_normalizzati_validi = righe_totali - valori_normalizzati_nulli
            
            # Conta i valori unici normalizzati (escludendo None)
            valori_unici_normalizzati = df_result['idItem_normalized'].nunique()
            
            self.log(f"📊 STATISTICHE NORMALIZZAZIONE:", "info")
            self.log(f"   Righe totali: {righe_totali}", "info")
            self.log(f"   Valori originali nulli: {valori_originali_nulli}", "info")
            self.log(f"   Valori originali vuoti: {valori_originali_vuoti}", "info")
            self.log(f"   Valori normalizzati validi: {valori_normalizzati_validi}", "info")
            self.log(f"   Valori normalizzati nulli: {valori_normalizzati_nulli}", "info")
            self.log(f"   Valori unici normalizzati: {valori_unici_normalizzati}", "info")
            
            # Verifica se ci sono errori di conversione
            errori_conversione = df_result[
                df_result['idItem'].notna() & 
                (df_result['idItem'] != '') & 
                df_result['idItem_normalized'].isna()
            ]
            
            if len(errori_conversione) > 0:
                self.log(f"⚠️  {len(errori_conversione)} valori non sono stati convertiti correttamente", "warning")
            
            self.log(f"✅ Normalizzazione completata con successo", "success")
            self.log(f"   DataFrame risultante: {len(df_result)} righe, {len(df_result.columns)} colonne", "success")
            
            return True, df_result
            
        except Exception as e:
            self.log(f"Errore nella normalizzazione: {str(e)}", "error")
            logger.error(f"Dettaglio errore: {str(e)}", exc_info=True)
            return False, None
    
    def extract_adm(self, df: pd.DataFrame) -> Tuple[bool, Optional[pd.DataFrame]]:
        """
        Estrae gli Avvisi di Manutenzione (AdM) dal DataFrame normalizzato.
        
        Args:
            df (pd.DataFrame): DataFrame normalizzato.
            
        Returns:
            tuple: (bool, DataFrame) - True se l'estrazione è riuscita, DataFrame degli AdM.
        """
        try:
            self.log(f"Estrazione AdM - Presenti {len(df)} idItem nel DataFrame")
            
            # Assicurati che i valori siano numerici
            try:
                # Crea nuovo DataFrame con valori AdM non-None, senza duplicati
                df_adm_unique = df[df['AdM'].notna()]['AdM'].drop_duplicates().astype(int).reset_index(drop=True).to_frame()
            except Exception as e:
                self.log(f"Errore nella estrazione della colonna 'AdM': {str(e)}", "error")
                return False, None

            
            self.log(f"Filtrati {len(df_adm_unique)} record con idItem < 2000000000 su {len(df)} totali", "success")
            
            return True, df_adm_unique
            
        except Exception as e:
            self.log(f"Errore nell'estrazione degli AdM: {str(e)}", "error")
            logger.error(f"Dettaglio errore: {str(e)}", exc_info=True)
            return False, None
    
    def extract_odm(self, df: pd.DataFrame) -> Tuple[bool, Optional[pd.DataFrame]]:
        """
        Estrae gli Ordini di Manutenzione (OdM) dal DataFrame normalizzato.
        
        Args:
            df (pd.DataFrame): DataFrame normalizzato.
            
        Returns:
            tuple: (bool, DataFrame) - True se l'estrazione è riuscita, DataFrame degli OdM.
        """
        try:
            self.log(f"Estrazione OdM - Presenti {len(df)} idItem nel DataFrame")
            
            # Assicurati che i valori siano numerici
            try:
                # Crea nuovo DataFrame con valori OdM non-None, senza duplicati
                df_OdM_unique = df[df['OdM'].notna()]['OdM'].drop_duplicates().astype(int).reset_index(drop=True).to_frame()
            except Exception as e:
                self.log(f"Errore nella estrazione della colonna 'OdM': {str(e)}", "error")
                return False, None

            
            self.log(f"Filtrati {len(df_OdM_unique)} record con idItem < 2000000000 su {len(df)} totali", "success")
            
            return True, df_OdM_unique
            
        except Exception as e:
            self.log(f"Errore nell'estrazione degli OdM: {str(e)}", "error")
            logger.error(f"Dettaglio errore: {str(e)}", exc_info=True)
            return False, None
    
 
    def get_df_excel_normalized(self) -> Optional[pd.DataFrame]:
        """
        Restituisce il DataFrame normalizzato contenente gli idItem.
        
        Returns:
            pd.DataFrame: DataFrame normalizzato con la colonna 'idItem_normalized'.
        """
        return self.df_excel_normalized
    
    def get_df_excel(self) -> Optional[pd.DataFrame]:
        """
        Restituisce il DataFrame contenente il file excel.
        
        Returns:
            pd.DataFrame: DataFrame contenente il file excel
        """
        return self.df_excel    

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
            "totale_righe_originali": len(self.df_excel_normalized) if self.df_excel_normalized is not None else 0,
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
        self.df_excel_normalized = None
        self.excel_file_path = None
        
        self.log("Dati elaborati cancellati", "info")