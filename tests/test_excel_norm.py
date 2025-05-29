#!/usr/bin/env python3
"""
Programma di test per la funzione normalize_df.
Carica un file Excel, applica la normalizzazione della colonna idItem,
e salva il risultato in un nuovo file Excel.
"""

import pandas as pd
import logging
from pathlib import Path
from typing import Tuple, Optional
import sys

# Configurazione del logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.StreamHandler(sys.stdout),
        logging.FileHandler('normalize_test.log', encoding='utf-8')
    ]
)
logger = logging.getLogger(__name__)


class DataNormalizer:
    """Classe per la normalizzazione dei dati"""
    
    def __init__(self):
        self.logger = logger
    
    def log(self, message: str, level: str = "info"):
        """Helper per il logging con diversi livelli"""
        if level == "error":
            self.logger.error(message)
        elif level == "warning":
            self.logger.warning(message)
        elif level == "success":
            self.logger.info(f"✅ {message}")
        else:
            self.logger.info(message)

    def normalize_df(self, df: pd.DataFrame) -> Tuple[bool, Optional[pd.DataFrame]]:
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
            
            # Rimuove eventuali righe dal DF in cui idItem è NaN o vuoto
            num_righe = len(df_result)
            df_not_nan = df_result[df_result['idItem'].notna() & (df_result['idItem'].str.strip() != '')]
            if num_righe != len(df_not_nan):
                self.log(f"Rimosse {num_righe - len(df_not_nan)} righe con idItem vuoto o NaN", "warning")
            df_result = df_not_nan
            self.log(f"Righe rimanenti dopo rimozione: {len(df_result)}", "info")


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
            
            # Mostra alcuni esempi di normalizzazione
            self.log("📋 ESEMPI DI NORMALIZZAZIONE:", "info")
            
            # Trova esempi di diverse tipologie
            esempi_mostrati = 0
            max_esempi = 5
            
            for idx, row in df_result.iterrows():
                if esempi_mostrati >= max_esempi:
                    break
                    
                id_originale = row['idItem']
                id_normalizzato = row['idItem_normalized']
                
                # Mostra esempi significativi
                if pd.notna(id_originale) and str(id_originale).strip():
                    if pd.notna(id_normalizzato):
                        self.log(f"   '{id_originale}' → {id_normalizzato}", "info")
                    else:
                        self.log(f"   '{id_originale}' → None (non convertibile)", "warning")
                    esempi_mostrati += 1
            
            # Verifica se ci sono errori di conversione
            errori_conversione = df_result[
                df_result['idItem'].notna() & 
                (df_result['idItem'] != '') & 
                df_result['idItem_normalized'].isna()
            ]
            
            if len(errori_conversione) > 0:
                self.log(f"⚠️  {len(errori_conversione)} valori non sono stati convertiti correttamente", "warning")
                # Mostra alcuni esempi di errori
                esempi_errori = errori_conversione['idItem'].head(3).tolist()
                self.log(f"   Esempi: {esempi_errori}", "warning")
            
            self.log(f"✅ Normalizzazione completata con successo", "success")
            self.log(f"   DataFrame risultante: {len(df_result)} righe, {len(df_result.columns)} colonne", "success")
            
            return True, df_result
            
        except Exception as e:
            self.log(f"Errore nella normalizzazione: {str(e)}", "error")
            logger.error(f"Dettaglio errore: {str(e)}", exc_info=True)
            return False, None


def main():
    """Funzione principale per il test della normalizzazione"""
    
    # Percorsi dei file
    input_file = r"C:\Users\a259046\OneDrive - Enel Spa\SCRIPT AHK e VBA\GITHUB\KPI_OFA\data\attivita_utenti_dal_01-04-25_al_29-04-25.xlsx"
    output_file = r"C:\Users\a259046\OneDrive - Enel Spa\SCRIPT AHK e VBA\GITHUB\KPI_OFA\data\attivita_utenti_dal_01-04-25_al_29-04-25_norm.xlsx"
    
    # Verifica esistenza file di input
    input_path = Path(input_file)
    if not input_path.exists():
        logger.error(f"❌ File di input non trovato: {input_file}")
        return False
    
    # Crea directory di output se non esiste
    output_path = Path(output_file)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    
    try:
        logger.info("🚀 INIZIO TEST NORMALIZZAZIONE")
        logger.info(f"📂 File input: {input_file}")
        logger.info(f"📂 File output: {output_file}")
        
        # Carica il file Excel
        logger.info("📖 Caricamento file Excel...")
        df = pd.read_excel(input_file)
        
        logger.info(f"✅ File caricato con successo: {len(df)} righe, {len(df.columns)} colonne")
        logger.info(f"📋 Colonne disponibili: {list(df.columns)}")
        
        # Mostra informazioni sulla colonna idItem se presente
        if 'idItem' in df.columns:
            logger.info("🔍 ANALISI COLONNA idItem:")
            logger.info(f"   Valori totali: {len(df)}")
            logger.info(f"   Valori nulli: {df['idItem'].isna().sum()}")
            logger.info(f"   Valori vuoti: {(df['idItem'] == '').sum()}")
            logger.info(f"   Valori unici: {df['idItem'].nunique()}")
            
            # Mostra alcuni esempi
            esempi = df['idItem'].dropna().head(10).tolist()
            logger.info(f"   Primi esempi: {esempi}")
        else:
            logger.warning("⚠️  Colonna 'idItem' non trovata nel DataFrame")
        
        # Inizializza il normalizzatore e applica la funzione
        normalizer = DataNormalizer()
        success, df_normalized = normalizer.normalize_df(df)
        
        if not success or df_normalized is None:
            logger.error("❌ Normalizzazione fallita")
            return False
        
        # Salva il risultato
        logger.info("💾 Salvataggio file Excel normalizzato...")
        df_normalized.to_excel(output_file, index=False)
        
        logger.info("✅ FILE SALVATO CON SUCCESSO!")
        logger.info(f"📊 RIEPILOGO FINALE:")
        logger.info(f"   Righe elaborate: {len(df_normalized)}")
        logger.info(f"   Colonne totali: {len(df_normalized.columns)}")
        logger.info(f"   File output: {output_file}")
        
        # Verifica finale del file salvato
        logger.info("🔍 Verifica file salvato...")
        df_check = pd.read_excel(output_file)
        logger.info(f"✅ Verifica completata: {len(df_check)} righe caricate dal file salvato")
        
        return True
        
    except FileNotFoundError as e:
        logger.error(f"❌ File non trovato: {e}")
        return False
    except pd.errors.EmptyDataError:
        logger.error("❌ Il file Excel è vuoto")
        return False
    except Exception as e:
        logger.error(f"❌ Errore durante l'elaborazione: {str(e)}")
        logger.error("Dettagli completi dell'errore:", exc_info=True)
        return False


if __name__ == "__main__":
    success = main()
    if success:
        logger.info("🎉 Test completato con successo!")
        sys.exit(0)
    else:
        logger.error("💥 Test fallito!")
        sys.exit(1)