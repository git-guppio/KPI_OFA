import pandas as pd
import os
from pathlib import Path
import logging

# Configurazione logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

class ExcelTestLoader:
    """Classe per caricare i file Excel di test e creare i dataframe corrispondenti"""
    
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
            "IW39_OdM_test.xlsx": "df_IW39"
        }
        self.dataframes = {}
    
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
    
    def load_excel_file(self, filename: str) -> pd.DataFrame:
        """
        Carica un singolo file Excel
        
        Args:
            filename (str): Nome del file da caricare
            
        Returns:
            pd.DataFrame: DataFrame caricato o DataFrame vuoto in caso di errore
        """
        file_path = self.base_path / filename
        
        try:
            # Prova a caricare il file Excel
            df = pd.read_excel(
                file_path,
                na_values=['', ' ', 'NULL', 'N/A'],  # Valori da considerare NaN
                header=0  # Modifica se header non è sulla prima riga
            )
            logger.info(f"File {filename} caricato con successo - Shape: {df.shape}")
            return df
            
        except FileNotFoundError:
            logger.error(f"File non trovato: {filename}")
            return pd.DataFrame()
            
        except Exception as e:
            logger.error(f"Errore durante il caricamento di {filename}: {str(e)}")
            return pd.DataFrame()
    
    def load_all_files(self) -> dict:
        """
        Carica tutti i file Excel e crea i dataframe corrispondenti
        
        Returns:
            dict: Dizionario con nome_df: DataFrame
        """
        if not self.validate_path():
            return {}
        
        file_status = self.check_files_exist()
        
        for filename, df_name in self.test_file_to_df_mapping.items():
            if file_status.get(filename, False):
                df = self.load_excel_file(filename)
                if not df.empty:
                    # Converti i valori del DataFrame secondo le regole definite e assegnalo al dizionario
                    self.dataframes[df_name] = self.converti_valori(df)
                    logger.info(f"DataFrame {df_name} creato con successo")
                else:
                    logger.warning(f"DataFrame {df_name} è vuoto")
            else:
                logger.warning(f"Saltato {filename} - file non trovato")
        
        return self.dataframes

    def converti_valori(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        Elabora un DataFrame convertendo ogni colonna secondo criteri specifici.

        La funzione analizza il contenuto di ogni colonna e applica le seguenti regole:
        a) Colonne di stringhe rimangono stringhe.
        b) Colonne di numeri interi vengono convertite in stringhe.
        c) Colonne di numeri float vengono prima convertite in interi (troncando i decimali)
        e poi in stringhe.
        d) Colonne con valori nel formato 'gg.mm.aaaa' vengono convertite in formato datetime.
        
        I valori mancanti (NaN) vengono preservati in tutte le conversioni.

        Args:
            df: Il DataFrame di input da elaborare.

        Returns:
            Un nuovo DataFrame con le colonne elaborate secondo le regole.
        """
        df_elaborato = df.copy()

        for colonna in df_elaborato.columns:
            if df_elaborato[colonna].isnull().all():
                continue
            date_convertite = None  # inizializzo la variabile per le date convertite
            # --- d) Criterio Data (logica aggiornata) ---
            if pd.api.types.is_object_dtype(df_elaborato[colonna].dtype):
                if df_elaborato[colonna].str.match(r'^\d{1,2}\.\d{1,2}\.\d{4}$').any():
                    date_convertite = pd.to_datetime(df_elaborato[colonna], format='%d.%m.%Y', errors='coerce')
                elif df_elaborato[colonna].str.match(r'^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}\.\d{3}Z$').any():
                    # Conversione per formato ISO 8601 con millisecondi
                    date_convertite = pd.to_datetime(df_elaborato[colonna], errors='coerce')
                # verifico se sono state converite alcune date
                if date_convertite is not None and date_convertite.notna().any():
                    # Verifica che TUTTI i valori non-missing siano stati convertiti.
                    numero_valori_originali = df_elaborato[colonna].notna().sum()
                    numero_valori_convertiti = date_convertite.notna().sum()
                    print(f"Colonna '{colonna}': Valori originali: {numero_valori_originali}, Valori convertiti: {numero_valori_convertiti}")
                    if numero_valori_originali > 0 and numero_valori_originali == numero_valori_convertiti:
                        df_elaborato[colonna] = date_convertite
                        print(f"Colonna '{colonna}': Tutti i valori convertiti in datetime.")
                    elif(numero_valori_convertiti > 0): # mostra i valori che non sono stati convertiti
                        valori_non_convertiti = df_elaborato[colonna][date_convertite.isna()]
                        print(f"Colonna '{colonna}': Valori non convertiti in datetime: {valori_non_convertiti.tolist()}")
                continue

            # --- c) Criterio Float -> Intero -> Stringa ---
            if pd.api.types.is_float_dtype(df_elaborato[colonna].dtype):
                df_elaborato[colonna] = df_elaborato[colonna].apply(
                    lambda x: str(int(x)).strip() if pd.notna(x) else None
                )
                continue

            # --- b) Criterio Intero -> Stringa ---
            if pd.api.types.is_integer_dtype(df_elaborato[colonna].dtype):
                df_elaborato[colonna] = df_elaborato[colonna].apply(
                    lambda x: str(x).strip() if pd.notna(x) else None
                )
                continue
                
            # --- a) Criterio Stringa -> Stringa ---
            if pd.api.types.is_object_dtype(df_elaborato[colonna].dtype):
                df_elaborato[colonna] = df_elaborato[colonna].apply(
                    lambda x: str(x).strip() if pd.notna(x) else None
                )

        return df_elaborato

    def get_dataframe_info(self) -> None:
        """Stampa informazioni sui dataframe caricati"""
        if not self.dataframes:
            logger.info("Nessun dataframe caricato")
            return
        
        print("\n" + "="*50)
        print("INFORMAZIONI SUI DATAFRAME CARICATI")
        print("="*50)
        
        for df_name, df in self.dataframes.items():
            print(f"\n{df_name}:")
            print(f"  - Shape: {df.shape}")
            print(f"  - Colonne: {list(df.columns)}")
            print(f"  - Tipi di dati: {df.dtypes.to_dict()}")
            print(f"  - Valori nulli: {df.isnull().sum().sum()}")

def process_iw29_with_ofa_actions(df_IW29, df_excel_normalized):
    """
    Processa il DataFrame df_IW29 aggiungendo colonne OFA basate sui dati di df_excel_normalized
    
    Args:
        df_IW29 (pd.DataFrame): DataFrame IW29 con colonne 'Avviso' e 'Ordine'
        df_excel_normalized (pd.DataFrame): DataFrame normalizzato con colonne 'AdM', 'OdM' e 'action'
        df_IW39 (pd.DataFrame): DataFrame IW39 con colonne 'Ordine' e 'Stato sistema'
    
    Returns:
        pd.DataFrame: DataFrame df_IW29 aggiornato con le nuove colonne OFA
        
    ELABORAZIONE NOTIFICATION:
    - CREATE NOTIFICATION -> Creati_in_OFA
    - DETAIL NOTIFICATION -> Visualizzati_in_OFA
    - UPDATE NOTIFICATION -> Modificati_in_OFA
    - OPEN TEXT UPLOAD ATTACHMEMT NOTIFICATION -> Allegati_in_OFA
    - OPEN TEXT UPLOAD ATTACHMEMT -> Allegati_in_OFA
    - CLOSED NOTIFICATION -> Chiusi_in_OFA
    - WORKORDER UPDATE TECO -> ricerca tramite Ordine in df_IW39 -> se Stato sistema contiene 'TECO' o 'CONC' -> Chiusi_in_OFA
    """
    
    # Verifica input
    if df_IW29 is None or df_IW29.empty:
        logger.error("df_IW29 è vuoto o None")
        return pd.DataFrame()
    
    if df_excel_normalized is None or df_excel_normalized.empty:
        logger.error("df_excel_normalized è vuoto o None")
        return df_IW29.copy()
    
    # Verifica colonne necessarie
    if 'Avviso' not in df_IW29.columns:
        logger.error("Colonna 'Avviso' non trovata in df_IW29")
        return df_IW29.copy()
    
    if 'Ordine' not in df_IW29.columns:
        logger.error("Colonna 'Ordine' non trovata in df_IW29")
        return df_IW29.copy()
    
    if 'AdM' not in df_excel_normalized.columns or 'action' not in df_excel_normalized.columns:
        logger.error("Colonne 'AdM' o 'action' non trovate in df_excel_normalized")
        return df_IW29.copy()
        
    # Crea una copia del DataFrame per non modificare l'originale
    df_result = df_IW29.copy()
    
    # Definisci le nuove colonne OFA
    ofa_columns = ['Creati_in_OFA', 'Visualizzati_in_OFA', 'Modificati_in_OFA', 
                   'Allegati_in_OFA', 'Chiusi_in_OFA']
    
    # Inizializza le nuove colonne con False
    for col in ofa_columns:
        df_result[col] = False
    
    # Mapping delle azioni alle colonne
    action_mapping = {
        'CREATE NOTIFICATION': 'Creati_in_OFA',
        'DETAIL NOTIFICATION': 'Visualizzati_in_OFA',
        'UPDATE NOTIFICATION': 'Modificati_in_OFA',
        'OPEN TEXT UPLOAD ATTACHMEMT NOTIFICATION': 'Allegati_in_OFA',
        'OPEN TEXT UPLOAD ATTACHMEMT': 'Allegati_in_OFA',
        'CLOSED NOTIFICATION': 'Chiusi_in_OFA',
        'WORKORDER UPDATE TECO': 'Chiusi_in_OFA'
    }

    # Creo un dizionario con le mascherre per le diverse azioni
    mask_dict = {}
    for action, target_column in action_mapping.items():
        df_action = df_excel_normalized[df_excel_normalized['action'] == action] # Creo un df filtrando df_excel_normalized in base all'azione considerata
        print(f"Dimensioni del df per action: {action} = {df_action.shape[0]}")
        if df_action.empty:
            logger.warning(f"Nessun AdM per azione '{action}'")
            mask_dict[action] = pd.Series([False] * len(df_result), index=df_result.index)  # Mask vuota
            continue
        if action == 'WORKORDER UPDATE TECO':
            # Verifica che ci siano OdM validi
            valid_odm = df_action['OdM'].dropna()
            print(f"Numero di OdM = {len(valid_odm)}")
            if valid_odm.empty:
                logger.warning(f"Nessun OdM valido per azione '{action}'")
                mask_dict[action] = pd.Series([False] * len(df_result), index=df_result.index)
                continue
            mask = df_result['Ordine'].isin(valid_odm)
            mask_dict[action] = mask
            # stampo il numero di elementi
            count1 = mask.sum()
            print(f"Mask - 'WORKORDER UPDATE TECO': {count1}") # -> OK

        else:    
            # Verifica che ci siano AdM validi
            valid_adm = df_action['AdM'].dropna()
            if valid_adm.empty:
                logger.warning(f"Nessun AdM valido per azione '{action}'")
                mask_dict[action] = pd.Series([False] * len(df_result), index=df_result.index)
                continue
                # Mask base
            mask = df_result['Avviso'].isin(valid_adm)
            mask_dict[action] = mask
    
    # Processa ogni colonna del df_result realizzando e applicando le maschere  secondo le azioni in OFA
    for target_column in ofa_columns:
        # Imposto un valore di default da utilizzare in caso di errore
        mask_finale = pd.Series([False] * len(df_result), index=df_result.index)

        # Condizioni speciali
        if target_column == 'Creati_in_OFA':
            mask_base = mask_dict['CREATE NOTIFICATION']
            # Aggiunge condizione Sis.Legacy
            if 'Sis.Legacy' in df_result.columns:
                mask_legacy = df_result['Sis.Legacy'].str.contains('OFA', na=False, case=False) # da aggiungere filtro sulla data creazione deve essere nel mese corrente
            else:
                print(f"ERRORE nella valutazione della colonna 'Sis.Legacy'")
                continue

            # Aggiungi condizione per verificare che l'AdM abbia il campo 'data_creazione' entro il range di date selezionato nella GUI
            if 'Data' in df_result.columns:
                # Definisci il range di date
                data_inizio = pd.to_datetime('01.05.2025', format='%d.%m.%Y')  # Le tue date
                data_fine = pd.to_datetime('31.05.2025', format='%d.%m.%Y')
                
                # Maschera per il range di date (estremi compresi) - NO conversione necessaria
                mask_date_range = (df_result['Data'] >= data_inizio) & \
                                (df_result['Data'] <= data_fine)
            else:
                print(f"ERRORE nella valutazione della colonna 'Data'")
                continue             
            
            ### Maschera con date
            # mask_finale = mask_date_range & (mask_base | mask_legacy)
            mask_finale = (mask_base | mask_legacy)
                
        elif target_column == 'Chiusi_in_OFA':
            # Condizioni
            mask_odm_teco = mask_dict['WORKORDER UPDATE TECO']
            mask_adm_closed = mask_dict['CLOSED NOTIFICATION']
            
            # Condizione valore ordine non vuoto
            if 'Ordine' in df_result.columns:
                mask_OdM = (
                    df_result['Ordine'].notna() &  # Non è NaN/None
                    (df_result['Ordine'].astype(str).str.strip() != '')  # Non è stringa vuota
                )
                count = mask_OdM.sum()
                print(f"Mask OdM non nulli: {count}") # -> OK
            else:
                print(f"ERRORE nella valutazione della colonna 'Ordine'")
                continue                

            # Condizione Avviso chiuso
            if 'St.sist.' in df_result.columns:
                mask_meco = df_result['St.sist.'].str.contains('MECO', na=False, case=False)
                count = mask_meco.sum()   
            else:
                print(f"ERRORE nella valutazione della colonna 'St.sist.'")
                continue                          

            mask_finale = mask_adm_closed | (mask_OdM & mask_odm_teco)
            ### da valutare se inserire controllo su AdM realmente chiusi
            # mask_finale = (mask_meco & mask_adm_closed) | (mask_OdM & mask_odm_teco)

        elif target_column == 'Allegati_in_OFA':
            # Condizioni
            mask_adm_upload_1 = mask_dict['OPEN TEXT UPLOAD ATTACHMEMT']
            mask_adm_upload_2 = mask_dict['OPEN TEXT UPLOAD ATTACHMEMT NOTIFICATION']
  
            mask_finale = mask_adm_upload_1 | mask_adm_upload_2

        elif target_column == 'Visualizzati_in_OFA':
            mask_finale = mask_dict['DETAIL NOTIFICATION']

        elif target_column == 'Modificati_in_OFA':
            mask_finale = mask_dict['UPDATE NOTIFICATION']           

        else:
            print(f"ERRORE nella valutazione della colonna del df")
        
        # Applica
        df_result.loc[mask_finale, target_column] = True
        logger.info(f"Colonna {target_column}: {df_result[target_column].sum()} True") # -> OK

    # Costruisce la colonne 'Numeratore'
    # Il numeratore assume valore 1 se almeno una delle colonne che descrivono gli stati OFA contiene un valore = 1
    # In pratica tutte le colonne sono in OR
    df_result['Numeratore'] = (df_result['Creati_in_OFA'] | 
                                df_result['Visualizzati_in_OFA'] |
                                df_result['Modificati_in_OFA'] |
                                df_result['Allegati_in_OFA'] |
                                df_result['Chiusi_in_OFA'])
    
    logger.info(f"Colonna {'Numeratore'}: {df_result['Numeratore'].sum()} True") # -> OK
    
    # Costruisco la colonna 'Denominatore'
    # Il denominatore assume valore 1 se il denominatore è 1 oppure se
    # la colonna 'Ordine' != "" E 
    # la colonna data modifica 'Mod. il' è vuota E
    # la data di inizio cardine è contenuta nel range di date selezionato nella GUI
    # altrimenti = 0
    
    # Costruisco la maschera a partire dai valori del numeratore
    mask_numeratore = df_result['Numeratore']

    # Definisco condizione per verificare che l'AdM abbia il campo 'OdM_data_inizio_cardine' entro il range di date selezionato nella GUI
    if 'OdM_data_inizio_cardine' in df_result.columns:
        # Definisci il range di date
        data_inizio = pd.to_datetime('01.05.2025', format='%d.%m.%Y')  # Le tue date
        data_fine = pd.to_datetime('31.05.2025', format='%d.%m.%Y')
        
        # Maschera per il range di date (estremi compresi)
        mask_date_range = (df_result['OdM_data_inizio_cardine'] >= data_inizio) & (df_result['OdM_data_inizio_cardine'] <= data_fine)
    else:
        print(f"ERRORE nella valutazione della colonna 'Data'")
        return False
    
    # Condizione valore ordine non vuoto
    if 'Ordine' in df_result.columns:
        mask_OdM = (
            df_result['Ordine'].notna() &  # Non è NaN/None
            (df_result['Ordine'].astype(str).str.strip() != '')  # Non è stringa vuota
        )
    else:
        print(f"ERRORE nella valutazione della colonna 'Ordine'")
        return False

    # Condizione colonna 'Mod. il' vuota 
    if 'Mod. il' in df_result.columns:
        mask_data_modifica = df_result['Mod. il'].notna()
    else:
        print(f"ERRORE nella valutazione della colonna 'Ordine'")
        return False

    df_result['Denominatore'] = False
    mask_finale = mask_numeratore | (mask_OdM & mask_data_modifica & mask_date_range)
    target_column = 'Denominatore'
    # Applica
    df_result.loc[mask_finale, target_column] = True
    logger.info(f"Colonna {target_column}: {df_result[target_column].sum()} True") # -> OK

    
    return df_result

def main():
    """Funzione principale per testare il caricamento dei file"""
    
    # Percorso dei file di test
    test_path = r"C:\Users\a259046\OneDrive - Enel Spa\SCRIPT AHK e VBA\GITHUB\KPI_OFA\data\test"
    
    # Crea l'istanza del loader
    loader = ExcelTestLoader(test_path)
    
    print("Inizio caricamento file Excel di test...")
    
    # Carica tutti i file
    dataframes = loader.load_all_files()
    
    # Mostra informazioni sui dataframe
    loader.get_dataframe_info()
    
    # Estrai i singoli dataframe per passarli alla funzione di elaborazione
    df_AFKO = dataframes.get('df_AFKO')
    df_excel_normalized = dataframes.get('df_excel_normalized')
    df_IW29 = dataframes.get('df_IW29')
    df_IW39 = dataframes.get('df_IW39')

    # Aggiungo la colonna 'OdM_data_inizio_cardine' nel df_IW29 ricavandolo dal df_AFKO
    # Crea dizionario di lookup per una singola colonna
    lookup_dict = df_AFKO.set_index('Ordine')['Data inizio cardine'].to_dict()
    df_IW29['OdM_data_inizio_cardine'] = df_IW29['Ordine'].map(lookup_dict)

    # Aggiungo la colonna 'Stato_sistema' nel df_IW29 ricavandolo dal df_IW39
    # Crea dizionario di lookup per una singola colonna
    lookup_dict = df_IW39.set_index('Ordine')['Stato sistema'].to_dict()
    df_IW29['OdM_stato_sistema'] = df_IW29['Ordine'].map(lookup_dict)    

    # Aggiungo la colonna 'Tecnologia'
    # Prelevo il terzo carattere della colonna 'Sede tecnica'
    df_IW29['Tecnologia'] = df_IW29['Sede tecnica'].str[2]  # Indice 2 = terzo carattere

    # Avvio l'elaborazione degli avvisi
    df = process_iw29_with_ofa_actions(df_IW29, df_excel_normalized)
    
    # Multiple funzioni di aggregazione
    pivot_multi = df.pivot_table(
        index=['Tecnologia', 'Pse'],  # Righe gerarchiche
        values=['Numeratore', 'Denominatore'],  # Colonna da sommare
        aggfunc='sum'  # Funzione di aggregazione
    ).assign(KPI=lambda x: x['Numeratore'] / x['Denominatore'] * 100)

    print(f"\nPivot con multiple aggregazioni:")
    print(pivot_multi.to_string())

    return dataframes


if __name__ == "__main__":
    # Esegui il test
    risultati = main()
    
    # I dataframe sono ora disponibili nel dizionario 'risultati'
    print(f"\nDataframe caricati: {list(risultati.keys())}")