import pandas as pd
import os
from pathlib import Path
import logging
import re

# Configurazione logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

class DfProcessor:
    """Classe per caricare i file Excel di test e creare i dataframe corrispondenti"""
    
    def __init__(self, df_dict: dict[str, pd.DataFrame], intervallo_date: tuple):

        """
        Inizializza i dataframe con i dati forniti
        
        Args:
            Dizionario con i nomi dei df come chiavi e i DataFrame come valori
        """
        self.df_dict = df_dict
        self.intervallo_date = intervallo_date

    def converti_tutti_df(self, dict_df: dict[str, pd.DataFrame])  -> tuple[bool, dict[str, pd.DataFrame] | None]:
        """
        Applica la conversione a tutti i DataFrame nel dizionario con gestione errori.
        
        Args:
            - dict_df: Dizionario con i DataFrame da convertire.
        Returns:
            Una tupla contenente:
            - bool: True se tutti i DataFrame sono stati elaborati con successo, False altrimenti.
            - dict: Dizionario con i DataFrame convertiti, o None in caso di errore.
            
        Raises:
            Exception: Se si verifica un errore durante l'elaborazione di un DataFrame.
        """
        df_dict_conv = {}
        errori = []
        df_num = 0
        df_count = 0
        try: 
            for nome_df, df in dict_df.items():
                df_num += 1
                try:
                    print(f"Elaborazione DataFrame: {nome_df}")
                    result, df_dict_conv[nome_df] = self.converti_valori_df(df)
                    if result:
                        print(f"DataFrame '{nome_df}' elaborato con successo.\n")
                        df_count += 1
                    else:
                        raise Exception(f"Errore nella conversione del DataFrame '{nome_df}'")
                except Exception as e:
                    errore_msg = f"Errore nell'elaborazione del DataFrame '{nome_df}': {str(e)}"
                    print(errore_msg)
                    errori.append(errore_msg)
        
            if errori:
                raise Exception(f"Errori durante l'elaborazione: {'; '.join(errori)}")
            
            if not df_dict_conv:
                raise Exception("Nessun DataFrame è stato elaborato correttamente.")
        
        except Exception as e:
            print(f"Errore durante l'elaborazione dei DataFrame: {str(e)}")
            return False, None
        
        print(f"\nElaborazione completata: {df_count}/{df_num} DataFrame elaborati con successo.")
        return True, df_dict_conv       


    def converti_valori_df(self, df: pd.DataFrame) -> tuple[bool, pd.DataFrame | None]:
        """
        Elabora un DataFrame convertendo ogni colonna secondo criteri specifici.

        La funzione analizza il contenuto di ogni colonna e applica le seguenti regole:
        1) Datetime con timezone vengono convertiti in datetime senza timezone
        2) Date italiane (gg.mm.aaaa) vengono convertite in formato datetime
        3) Colonne float vengono convertite in interi (troncando) e poi in stringhe
        4) Colonne integer vengono convertite in stringhe
        5) Colonne object/stringhe vengono pulite e mantenute come stringhe
        
        I valori mancanti (NaN) vengono preservati in tutte le conversioni.

        Args:
            df: Il DataFrame di input da elaborare.

        Returns:
            tuple: (success: bool, dataframe: pd.DataFrame | None)
            
        Raises:
            TypeError: Se l'input non è un DataFrame pandas.
            ValueError: Se il DataFrame è vuoto.
        """
        # Validazione input
        if not isinstance(df, pd.DataFrame):
            raise TypeError(f"L'input deve essere un DataFrame pandas, ricevuto: {type(df)}")
        
        if df.empty:
            raise ValueError("Il DataFrame fornito è vuoto")
        
        try:
            df_elaborato = df.copy()
            errori_colonne = []
            conversioni_log = []
            
            print(f"Inizio elaborazione DataFrame con {len(df_elaborato.columns)} colonne...")
            
            for colonna in df_elaborato.columns:
                try:
                    print(f"Elaborazione colonna: '{colonna}' (tipo: {df_elaborato[colonna].dtype})")
                    
                    # Salta colonne completamente vuote
                    if df_elaborato[colonna].isnull().all():
                        print(f"Colonna '{colonna}': completamente vuota, saltata.")
                        conversioni_log.append(f"{colonna}: SALTATA (vuota)")
                        continue
                    
                    # Conta valori non nulli per statistiche
                    valori_originali = df_elaborato[colonna].notna().sum()
                    
                    # PRIORITA' 1: Datetime già tipizzati con timezone
                    if pd.api.types.is_datetime64_any_dtype(df_elaborato[colonna]):
                        success = self._converti_datetime_tipizzato(df_elaborato, colonna)
                        if success:
                            conversioni_log.append(f"{colonna}: DATETIME con timezone -> DATETIME pulito ({valori_originali} valori)")
                            continue
                    
                    # PRIORITA' 2: Float -> Integer -> String
                    if pd.api.types.is_float_dtype(df_elaborato[colonna].dtype):
                        success = self._converti_float_to_string(df_elaborato, colonna)
                        if success:
                            conversioni_log.append(f"{colonna}: FLOAT -> STRING ({valori_originali} valori)")
                            continue
                        else:
                            errori_colonne.append(f"Errore conversione float->string in colonna '{colonna}'")
                    
                    # PRIORITA' 3: Integer -> String
                    if pd.api.types.is_integer_dtype(df_elaborato[colonna].dtype):
                        success = self._converti_integer_to_string(df_elaborato, colonna)
                        if success:
                            conversioni_log.append(f"{colonna}: INTEGER -> STRING ({valori_originali} valori)")
                            continue
                        else:
                            errori_colonne.append(f"Errore conversione integer->string in colonna '{colonna}'")
                    
                    # PRIORITA' 4: Object - Analisi contenuto
                    if pd.api.types.is_object_dtype(df_elaborato[colonna].dtype):
                        
                        # 4a) Datetime con timezone come stringhe
                        if self._ha_datetime_timezone(df_elaborato[colonna]):
                            valori_convertiti = self._converti_datetime_timezone_regex(df_elaborato, colonna)
                            conversioni_log.append(f"{colonna}: DATETIME timezone strings -> DATETIME ({valori_convertiti}/{valori_originali} convertiti)")
                            continue
                        
                        # 4b) Date italiane (gg.mm.aaaa)
                        if self._ha_date_italiane(df_elaborato[colonna]):
                            valori_convertiti = self._converti_date_italiane(df_elaborato, colonna)
                            conversioni_log.append(f"{colonna}: DATE italiane -> DATETIME ({valori_convertiti}/{valori_originali} convertiti)")
                            continue
                        
                        # 4c) Stringhe generiche - pulizia
                        success = self._converti_stringhe_generiche(df_elaborato, colonna)
                        if success:
                            conversioni_log.append(f"{colonna}: STRINGHE pulite ({valori_originali} valori)")
                            continue
                        else:
                            errori_colonne.append(f"Errore pulizia stringhe in colonna '{colonna}'")
                    
                    # Se arriviamo qui, tipo non gestito
                    print(f"ATTENZIONE: Colonna '{colonna}' con tipo '{df_elaborato[colonna].dtype}' non gestita")
                    conversioni_log.append(f"{colonna}: TIPO NON GESTITO ({df_elaborato[colonna].dtype})")
                    
                except Exception as e:
                    errore_msg = f"Errore generale nella colonna '{colonna}': {str(e)}"
                    errori_colonne.append(errore_msg)
                    print(f"ERRORE: {errore_msg}")
                    continue
            
            # Log finale delle conversioni
            print(f"\n=== RIEPILOGO CONVERSIONI ===")
            for log in conversioni_log:
                print(f"  {log}")
            
            # Gestione errori
            if errori_colonne:
                print(f"\n=== ERRORI RISCONTRATI ({len(errori_colonne)}) ===")
                for errore in errori_colonne:
                    print(f"  ❌ {errore}")
                # Decidi se proseguire o fermarsi
                if len(errori_colonne) > len(df_elaborato.columns) * 0.5:  # Se più del 50% ha errori
                    print("TROPPI ERRORI: Elaborazione interrotta")
                    return False, None
                else:
                    print("PROSEGUIMENTO: Errori parziali, DataFrame comunque utilizzabile")
            
            print(f"\nElaborazione completata: {df_elaborato.shape[0]} righe, {df_elaborato.shape[1]} colonne")
            return True, df_elaborato
            
        except Exception as e:
            print(f"ERRORE CRITICO durante l'elaborazione: {str(e)}")
            return False, None

    def _converti_datetime_tipizzato(self, df_elaborato, colonna):
        """Gestisce datetime già tipizzati rimuovendo timezone"""
        try:
            if hasattr(df_elaborato[colonna].dtype, 'tz') and df_elaborato[colonna].dtype.tz is not None:
                df_elaborato[colonna] = df_elaborato[colonna].dt.tz_localize(None)
                print(f"Colonna '{colonna}': Rimosso timezone da datetime tipizzato")
            elif df_elaborato[colonna].dt.tz is not None:
                df_elaborato[colonna] = df_elaborato[colonna].dt.tz_localize(None)
                print(f"Colonna '{colonna}': Rimosso timezone da datetime")
            return True
        except Exception as e:
            print(f"Errore rimozione timezone da '{colonna}': {str(e)}")
            return False

    def _converti_float_to_string(self, df_elaborato, colonna):
        """Converte float -> int -> string"""
        try:
            def converti_float_sicuro(x):
                if pd.notna(x):
                    if abs(x) > 1e15:
                        raise ValueError(f"Valore troppo grande: {x}")
                    return str(int(x)).strip()
                return None
            
            df_elaborato[colonna] = df_elaborato[colonna].apply(converti_float_sicuro)
            print(f"Colonna '{colonna}': Float -> String completata")
            return True
        except Exception as e:
            print(f"Errore conversione float '{colonna}': {str(e)}")
            return False

    def _converti_integer_to_string(self, df_elaborato, colonna):
        """Converte integer -> string"""
        try:
            df_elaborato[colonna] = df_elaborato[colonna].apply(
                lambda x: str(x).strip() if pd.notna(x) else None
            )
            print(f"Colonna '{colonna}': Integer -> String completata")
            return True
        except Exception as e:
            print(f"Errore conversione integer '{colonna}': {str(e)}")
            return False

    def _ha_datetime_timezone(self, series):
        """Verifica se la serie contiene datetime con timezone"""
        if series.empty:
            return False
        
        sample_values = series.dropna().astype(str).head(20)
        if sample_values.empty:
            return False
        
        timezone_patterns = [
            r'\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}[+\-]\d{2}:\d{2}',  # 2025-01-15T10:30:00+02:00
            r'\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}(?:\.\d{3})?Z',      # 2025-01-15T10:30:00.123Z
            r'\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2} [A-Z]{3,4}',        # 2025-01-15 10:30:00 UTC
            r'\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}[+\-]\d{4}'          # 2025-01-15 10:30:00+0200
        ]
        
        for pattern in timezone_patterns:
            if sample_values.str.contains(pattern, regex=True, na=False).any():
                return True
        return False

    def _converti_datetime_timezone_regex(self, df_elaborato, colonna):
        """Converte datetime con timezone usando regex"""
        patterns_with_groups = [
            r'(\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2})[+\-]\d{2}:\d{2}',
            r'(\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2})(?:\.\d{3})?Z',
            r'(\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}) [A-Z]{3,4}',
            r'(\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2})[+\-]\d{4}'
        ]
        
        valori_convertiti = 0
        
        def converti_valore(x):
            nonlocal valori_convertiti
            if pd.isna(x):
                return None
            
            str_val = str(x)
            for pattern in patterns_with_groups:
                match = re.match(pattern, str_val)
                if match:
                    try:
                        datetime_pulito = match.group(1)
                        dt = pd.to_datetime(datetime_pulito, errors='coerce')
                        if pd.notna(dt):
                            valori_convertiti += 1
                            return dt
                    except:
                        continue
            return x
        
        df_elaborato[colonna] = df_elaborato[colonna].apply(converti_valore)
        print(f"Colonna '{colonna}': Convertiti {valori_convertiti} datetime con timezone")
        return valori_convertiti

    def _ha_date_italiane(self, series):
        """Verifica se la serie contiene date italiane gg.mm.aaaa"""
        if series.empty:
            return False
        
        sample = series.dropna().astype(str).head(10)
        return sample.str.match(r'^\d{1,2}\.\d{1,2}\.\d{4}$', na=False).any()

    def _converti_date_italiane(self, df_elaborato, colonna):
        """Converte date italiane gg.mm.aaaa -> datetime"""
        try:
            date_convertite = pd.to_datetime(df_elaborato[colonna], format='%d.%m.%Y', errors='coerce')
            valori_originali = df_elaborato[colonna].notna().sum()
            valori_convertiti = date_convertite.notna().sum()
            
            if valori_convertiti > 0:
                df_elaborato[colonna] = date_convertite
                print(f"Colonna '{colonna}': Convertite {valori_convertiti}/{valori_originali} date italiane")
                return valori_convertiti
            return 0
        except Exception as e:
            print(f"Errore conversione date italiane '{colonna}': {str(e)}")
            return 0

    def _converti_stringhe_generiche(self, df_elaborato, colonna):
        """Pulisce e standardizza stringhe generiche"""
        try:
            def pulisci_stringa(x):
                if pd.notna(x):
                    str_value = str(x).strip()
                    if len(str_value) > 1000:  # Limite ragionevole
                        print(f"AVVISO: Stringa molto lunga in '{colonna}' troncata")
                        str_value = str_value[:1000]
                    return str_value if str_value else None
                return None
            
            df_elaborato[colonna] = df_elaborato[colonna].apply(pulisci_stringa)
            print(f"Colonna '{colonna}': Stringhe pulite e standardizzate")
            return True
        except Exception as e:
            print(f"Errore pulizia stringhe '{colonna}': {str(e)}")
            return False
    
    
    def converti_valori_df_old(self, df: pd.DataFrame)  -> tuple[bool, pd.DataFrame | None]:
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
        if not self.df_dict:
            logger.info("Nessun dataframe caricato")
            return
        
        print("\n" + "="*50)
        print("INFORMAZIONI SUI DATAFRAME CARICATI")
        print("="*50)
        
        for df_name, df in self.df_dict.items():
            print(f"\n{df_name}:")
            print(f"  - Shape: {df.shape}")
            print(f"  - Colonne: {list(df.columns)}")
            print(f"  - Tipi di dati: {df.dtypes.to_dict()}")
            print(f"  - Valori nulli: {df.isnull().sum().sum()}")

    def process_iw29_with_ofa_actions(self, df_IW29: pd.DataFrame, df_excel_normalized: pd.DataFrame) -> tuple[bool, pd.DataFrame | None]:
        """
        Processa il DataFrame df_IW29 aggiungendo colonne OFA basate sui dati di df_excel_normalized
        
        Args:
            df_IW29 (pd.DataFrame): DataFrame IW29 con colonne 'Avviso' e 'Ordine'
            df_excel_normalized (pd.DataFrame): DataFrame normalizzato con colonne 'AdM', 'OdM' e 'action'
        
        Returns:
            tuple[bool,pd.DataFrame: DataFrame df_IW29 aggiornato con le nuove colonne OFA]
            
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
            return False, None
        
        if df_excel_normalized is None or df_excel_normalized.empty:
            logger.error("df_excel_normalized è vuoto o None")
            return False, None
        
        # Verifica colonne necessarie
        if 'Avviso' not in df_IW29.columns:
            logger.error("Colonna 'Avviso' non trovata in df_IW29")
            return False, None
        
        if 'Ordine' not in df_IW29.columns:
            logger.error("Colonna 'Ordine' non trovata in df_IW29")
            return False, None
        
        if 'AdM' not in df_excel_normalized.columns or 'action' not in df_excel_normalized.columns:
            logger.error("Colonne 'AdM' o 'action' non trovate in df_excel_normalized")
            return False, None
            
        # Crea una copia del DataFrame per non modificare l'originale
        df_result = df_IW29.copy()
        
        # Definisci le nuove colonne OFA
        ofa_columns = [
            'Creati_in_OFA', 'Visualizzati_in_OFA', 'Modificati_in_OFA', 
            'Allegati_in_OFA', 'Chiusi_in_OFA'
        ]
        
        # Aggiungi e Inizializza le nuove colonne con False
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

        # Creo un dizionario con le maschere per le diverse azioni
        mask_dict = {}
        for action, target_column in action_mapping.items():
            try:
                df_action = df_excel_normalized[df_excel_normalized['action'] == action] # Creo un df filtrando df_excel_normalized in base all'azione considerata
                print(f"Dimensioni del df per action: {action} = {df_action.shape[0]}")
                if df_action.empty: # Se non esistono elementi per l'azione considerata esamino il successivo
                    logger.warning(f"Nessun AdM per azione '{action}'")
                    mask_dict[action] = pd.Series([False] * len(df_result), index=df_result.index)  # Mask vuota
                    continue
                if action == 'WORKORDER UPDATE TECO':
                    # Verifica che ci siano OdM validi
                    is_valid = (
                        df_action['OdM'].notna() & 
                        (df_action['OdM'].astype(str).str.strip() != '')
                    )
                    valid_odm = df_action.loc[is_valid, 'OdM']
                    
                    if valid_odm.empty:
                        logger.warning(f"Nessun OdM valido per azione '{action}'")
                        mask_dict[action] = pd.Series([False] * len(df_result), index=df_result.index)
                        continue

                    print(f"Numero di OdM validi = {len(valid_odm)}")
                    mask = df_result['Ordine'].isin(valid_odm)
                    mask_dict[action] = mask
                    # stampo il numero di elementi
                    count1 = mask.sum()
                    print(f"Mask - 'WORKORDER UPDATE TECO': {count1}") # -> OK

                else:    
                    # Verifica che ci siano AdM validi
                    is_valid = (
                        df_action['AdM'].notna() & 
                        (df_action['AdM'].astype(str).str.strip() != '')
                    )
                    valid_adm = df_action.loc[is_valid, 'AdM']
                    
                    if valid_adm.empty:
                        print(f"Nessun AdM valido per azione '{action}'")
                        mask_dict[action] = pd.Series([False] * len(df_result), index=df_result.index)
                        continue

                    print(f"Numero di AdM = {len(valid_adm)}")
                    mask = df_result['Avviso'].isin(valid_adm)
                    mask_dict[action] = mask

            except Exception as e:
                logger.error(f"IW29 Errore nella creazione della maschera per azione '{action}': {str(e)}")
                return False, None
            
        # inserire verifica numero elmenti del dizionario mask_dict con azioni in action_mapping
        print(f"Dizionario maschere creato con {len(mask_dict)} azioni.")
        if len(mask_dict) != len(action_mapping):
            logger.error("IW29 Il numero di maschere create non corrisponde al numero di azioni mappate.")
            return False, None

        # Definisce il range di date selezionato nella GUI
        pd_data_inizio = pd.to_datetime(self.intervallo_date[0].toString("dd.MM.yyyy"), format='%d.%m.%Y')  
        pd_data_fine = pd.to_datetime(self.intervallo_date[1].toString("dd.MM.yyyy"), format='%d.%m.%Y')                
        
        # Processa ogni colonna del df_result realizzando e applicando le maschere e i flitri sulle date secondo le azioni in OFA
        error_msg = ""
        for target_column in ofa_columns:
            try:
                # Imposto un valore di default da utilizzare in caso di errore
                mask_finale = pd.Series([False] * len(df_result), index=df_result.index)

                # Condizioni speciali
                if target_column == 'Creati_in_OFA':
                    mask_base = mask_dict['CREATE NOTIFICATION']
                    # Aggiunge condizione Sis.Legacy
                    if 'Sis.Legacy' in df_result.columns:
                        mask_legacy = df_result['Sis.Legacy'].str.contains('OFA', na=False, case=False)
                    else:
                        error_msg += "ERRORE nella valutazione della colonna 'Sis.Legacy'\n"
                        continue

                    # Verifico che l'AdM abbia il campo 'data_creazione' entro il range di date selezionato nella GUI
                    if 'Data cr.' in df_result.columns:                    
                        try:
                            # Converti solo se non è già datetime
                            if not pd.api.types.is_datetime64_any_dtype(df_result['Data cr.']):
                                df_result['Data cr.'] = pd.to_datetime(
                                    df_result['Data cr.'],
                                    format='%d.%m.%Y',
                                    errors='coerce'
                                )
                            # Creao Maschera per il range di date (estremi compresi)
                            mask_date_range = (
                                (df_result['Data cr.'] >= pd_data_inizio) &
                                (df_result['Data cr.'] <= pd_data_fine)
                            )
                        except Exception as e:
                            error_msg += f"ERRORE nella creazione della maschera date per 'Data cr.': {str(e)}\n"
                            continue                        
                    else:
                        error_msg += "ERRORE nella valutazione della colonna 'Data creazione'\n"
                        continue

                    ### Maschera con date
                    mask_finale = mask_date_range & (mask_base | mask_legacy)
                    #mask_finale = (mask_base | mask_legacy)
                        
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
                        error_msg += "ERRORE nella valutazione della colonna 'Ordine'\n"
                        continue                

                    # Condizione Avviso chiuso
                    if 'St.sist.' in df_result.columns:
                        mask_meco = df_result['St.sist.'].str.contains('MECO', na=False, case=False)
                        count = mask_meco.sum()
                        print(f"Mask AdM MECO: {count}") # -> OK 
                    else:
                        error_msg += "ERRORE nella valutazione della colonna 'St.sist.'\n"
                        continue                          

                    #mask_finale = mask_adm_closed | (mask_OdM & mask_odm_teco)
                    ### da valutare se inserire controllo su AdM realmente chiusi
                    mask_finale = (mask_meco & (mask_adm_closed | (mask_OdM & mask_odm_teco)))

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
                    error_msg += f"ERRORE nella valutazione della colonna: {target_column} del df\n"
                    continue # Se rilevo un errore non procedo con l'applicazione della maschera ma continuo l'iterazione

                # Applica
                df_result.loc[mask_finale, target_column] = True
                logger.info(f"Colonna {target_column}: {df_result[target_column].sum()} True") # -> OK
            
            except Exception as e:
                logger.error(f"Errore nell'applicazione della maschera alla colonna: '{target_column}': {str(e)}")
                return False, None
        
        if error_msg:
                logger.error(f"IW29 - Errori riscontrati: \n\t{error_msg}")
                return False, None

        # ========================================
        # Costruisce la colonne 'Numeratore'
        # ========================================        
        # Il numeratore assume valore 1 se almeno una delle colonne che descrivono gli stati OFA contiene un valore = 1
        # In pratica tutte le colonne sono in OR
        df_result['Numeratore'] = (
            df_result['Creati_in_OFA'] | 
            df_result['Visualizzati_in_OFA'] |
            df_result['Modificati_in_OFA'] |
            df_result['Allegati_in_OFA'] |
            df_result['Chiusi_in_OFA']
        )
        
        logger.info(f"Colonna {'Numeratore'}: {df_result['Numeratore'].sum()} True") # -> OK
        
        # ========================================
        # Costruisco la colonna 'Denominatore'
        # ========================================
        # Il denominatore assume valore 1 se il numeratore è 1 oppure se
        # la colonna 'Ordine' != "" E 
        # la colonna data modifica dell'avviso 'Mod. il' è contenuta nel range di date selezionato nella GUI E
        # la data di inizio cardine è contenuta nel range di date selezionato nella GUI
        # altrimenti = 0
        
        # Costruisco la maschera a partire dai valori del numeratore
        mask_numeratore = df_result['Numeratore']

        # Definisco condizione per verificare che l'AdM abbia il campo 'OdM_data_inizio_cardine' entro il range di date selezionato nella GUI
        if 'OdM_data_inizio_cardine' in df_result.columns:
            try:
                # Converti solo se non è già datetime
                if not pd.api.types.is_datetime64_any_dtype(df_result['OdM_data_inizio_cardine']):
                    df_result['OdM_data_inizio_cardine'] = pd.to_datetime(
                        df_result['OdM_data_inizio_cardine'],
                        format='%d.%m.%Y',
                        errors='coerce'
                    )
                # Maschera per il range di date (estremi compresi)
                mask_date_range = (
                    (df_result['OdM_data_inizio_cardine'] >= pd_data_inizio) &
                    (df_result['OdM_data_inizio_cardine'] <= pd_data_fine)
                )
            except Exception as e:
                error_msg += f"ERRORE nella creazione della maschera 'mask_date_range': {str(e)}\n"
                return False, None
        else:
            logger.error(f"ERRORE nella valutazione della colonna 'OdM_data_inizio_cardine'\n")
            return False, None
        
        # Condizione valore ordine non vuoto
        if 'Ordine' in df_result.columns:
            mask_OdM = (
                df_result['Ordine'].notna() &  # Non è NaN/None
                (df_result['Ordine'].astype(str).str.strip() != '')  # Non è stringa vuota
            )
        else:
            logger.error(f"ERRORE nella valutazione della colonna 'Ordine'")
            return False, None

        # Condizione colonna 'Mod. il' entro il range di date selezionato nella GUI
        if 'Mod. il' in df_result.columns:
            try:
                # Converti solo se non è già datetime
                if not pd.api.types.is_datetime64_any_dtype(df_result['Mod. il']):
                    df_result['Mod. il'] = pd.to_datetime(
                        df_result['Mod. il'],
                        format='%d.%m.%Y',
                        errors='coerce'
                    )
                # Maschera per il range di date (estremi compresi)
                mask_data_modifica = (
                    (df_result['Mod. il'] >= pd_data_inizio) &
                    (df_result['Mod. il'] <= pd_data_fine)
                )
            except Exception as e:
                logger.error(f"ERRORE nella creazione della maschera 'mask_data_modifica': {str(e)}\n")
                return False, None
        else:
            logger.error(f"ERRORE nella valutazione della colonna 'Mod. il'")
            return False, None

        df_result['Denominatore'] = False
        mask_finale = mask_numeratore | (mask_OdM & mask_data_modifica & mask_date_range)
        target_column = 'Denominatore'
        # Applica
        df_result.loc[mask_finale, target_column] = True
        logger.info(f"Colonna {target_column}: {df_result[target_column].sum()} True") # -> OK

        
        return True, df_result

    def process_iw39_with_ofa_actions(self, df_IW39, df_excel_normalized) -> tuple[bool, pd.DataFrame | None]:
        """
        Processa il DataFrame df_IW39 aggiungendo colonne OFA basate sui dati di df_excel_normalized
        
        Args:
            df_IW39 (pd.DataFrame): DataFrame IW39
            df_excel_normalized (pd.DataFrame): DataFrame normalizzato con colonne 'AdM', 'OdM' e 'action'
        
        Returns:
            tuple[bool,pd.DataFrame: DataFrame df_IW39 aggiornato con le nuove colonne OFA]
                        
            - Creati_in_OFA:
                Sis.Legacy = OFA -> 1 oppure:                      
            - Dal file di Simone cerca se l'OdM ha uno dei seguneti stati:
                CREATE WORKORDER -> Creati_in_OFA
                DETAIL WORK ORDER -> Visualizzati_in_OFA
                WORKORDER UPDATE RELEASE, WORKORDER HEADER UPDATE -> Modificati_in_OFA
                WORKORDER UPDATE OPERATION -> Operazione_gestita_in_OFA
                WORKORDER ADD COMPONENT -> Aggiunto_materiale
                WORKORDER UPDATE COMPONENT, DETAIL WORK ORDER MATERIAL, WORKORDER DELETE COMPONENT -> Gestione_materiale
                OPEN TEXT UPLOAD ATTACHMEMT WORKORDER, OPEN TEXT DETAIL WORK ORDER OPERATION -> Allegati_in_OFA
                DETAIL WORK ORDER CONFERMATION, CREATE TIME CONFERMATION -> Conferma_ore_in_OFA
                WORKORDER UPDATE TECO -> Chiusi_in_OFA
        """
        
        # Verifica input
        if df_IW39 is None or df_IW39.empty:
            logger.error("df_IW39 è vuoto o None")
            return False, None
        
        if df_excel_normalized is None or df_excel_normalized.empty:
            logger.error("df_excel_normalized è vuoto o None")
            return False, None
        
        if 'Ordine' not in df_IW39.columns:
            logger.error("Colonna 'Ordine' non trovata in df_IW39")
            return False, None
        
        if 'OdM' not in df_excel_normalized.columns or 'action' not in df_excel_normalized.columns:
            logger.error("Colonne 'OdM' o 'action' non trovate in df_excel_normalized")
            return False, None
            
        # Crea una copia del DataFrame per non modificare l'originale
        df_result = df_IW39.copy()
        
        # Definisci le nuove colonne OFA
        ofa_columns = [
            'Creati_in_OFA', 'Visualizzati_in_OFA', 'Modificati_in_OFA', 'Operazione_gestita_in_OFA',
            'Aggiunto_materiale', 'Gestione_materiale', 'Allegati_in_OFA', 'Conferma_ore_in_OFA',
            'Chiusi_in_OFA'
        ]
        
        # Inizializza le nuove colonne con False
        for col in ofa_columns:
            df_result[col] = False
        
        # Mapping delle azioni alle colonne
        action_mapping = {
            'CREATE WORKORDER': 'Creati_in_OFA',
            'DETAIL WORK ORDER': 'Visualizzati_in_OFA',
            'WORKORDER UPDATE RELEASE': 'Modificati_in_OFA',
            'WORKORDER HEADER UPDATE': 'Modificati_in_OFA',
            'WORKORDER UPDATE OPERATION': 'Operazione_gestita_in_OFA',
            'WORKORDER ADD COMPONENT': 'Aggiunto_materiale',
            'WORKORDER UPDATE COMPONENT': 'Gestione_materiale',
            'DETAIL WORK ORDER MATERIAL': 'Gestione_materiale',
            'WORKORDER DELETE COMPONENT': 'Gestione_materiale',
            'OPEN TEXT UPLOAD ATTACHMEMT WORKORDER': 'Allegati_in_OFA',
            'OPEN TEXT DETAIL WORK ORDER OPERATION': 'Allegati_in_OFA',
            'DETAIL WORK ORDER CONFERMATION': 'Conferma_ore_in_OFA',
            'CREATE TIME CONFERMATION': 'Conferma_ore_in_OFA',
            'WORKORDER UPDATE TECO': 'Chiusi_in_OFA'
        }

        # Creo un dizionario con le mascherre per le diverse azioni
        mask_dict = {}
        for action, target_column in action_mapping.items():
            try:
                # Creo un df filtrando df_excel_normalized in base all'azione considerata
                df_action = df_excel_normalized[df_excel_normalized['action'] == action] 
                print(f"Dimensioni del df per action: {action} = {df_action.shape[0]}")
                
                if df_action.empty: # Se non contiene valori allora creo una maschera solo False
                    logger.warning(f"Nessun OdM per azione '{action}'")
                    mask_dict[action] = pd.Series([False] * len(df_result), index=df_result.index)  # Mask vuota
                    continue

                # Verifica che ci siano OdM validi
                is_valid = (
                    df_action['OdM'].notna() & 
                    (df_action['OdM'].astype(str).str.strip() != '')
                )
                valid_odm = df_action.loc[is_valid, 'OdM']

                if valid_odm.empty: # Se non contiene valori allora creo una maschera solo False
                    logger.warning(f"Nessun OdM valido per azione '{action}'")
                    mask_dict[action] = pd.Series([False] * len(df_result), index=df_result.index)
                    continue
                # creo una maschera relativa all'azione considerata 
                mask = df_result['Ordine'].isin(valid_odm) # Verifico quali OdM in df_action sono presenti in df_result 
                mask_dict[action] = mask

            except Exception as e:
                logger.error(f"IW39 Errore nella creazione della maschera per azione '{action}': {str(e)}")
                return False, None

        # Verifica numero elmenti del dizionario mask_dict con azioni in action_mapping
        print(f"IW39 Dizionario maschere creato con {len(mask_dict)} azioni.")
        if len(mask_dict) != len(action_mapping):
            logger.error("IW39 Il numero di maschere create non corrisponde al numero di azioni mappate.")
            return False, None      

        # Definisce il range di date selezionato nella GUI
        pd_data_inizio = pd.to_datetime(self.intervallo_date[0].toString("dd.MM.yyyy"), format='%d.%m.%Y')  
        pd_data_fine = pd.to_datetime(self.intervallo_date[1].toString("dd.MM.yyyy"), format='%d.%m.%Y')                        

        # Processa ogni colonna del df_result realizzando e applicando le maschere  secondo le azioni in OFA
        error_msg = ""
        for target_column in ofa_columns:
            try:
                # Imposto un valore di default da utilizzare in caso di errore
                mask_finale = pd.Series([False] * len(df_result), index=df_result.index)

                # Condizioni speciali
                if target_column == 'Creati_in_OFA': # -> OK
                    mask_base = mask_dict['CREATE WORKORDER']
                    # Aggiunge condizione Sis.Legacy
                    if 'Sis Legacy' in df_result.columns:
                        mask_legacy = df_result['Sis Legacy'].str.contains('OFA', na=False, case=False) # da aggiungere filtro sulla data creazione deve essere nel mese corrente
                    else:
                        error_msg += "ERRORE nella valutazione della colonna 'Sis Legacy'"
                        continue

                    # Aggiungi condizione per verificare che l'OdM abbia il campo data di acquisizione 'Data acq.' entro il range di date selezionato nella GUI
                    if 'Data acq.' in df_result.columns:
                        try:
                            # Converti solo se non è già datetime
                            if not pd.api.types.is_datetime64_any_dtype(df_result['Data acq.']):
                                df_result['Data acq.'] = pd.to_datetime(
                                    df_result['Data acq.'],
                                    format='%d.%m.%Y',
                                    errors='coerce'
                                )
                            # Creao Maschera per il range di date (estremi compresi)
                            mask_date_range = (
                                (df_result['Data acq.'] >= pd_data_inizio) &
                                (df_result['Data acq.'] <= pd_data_fine)
                            )
                        except Exception as e:
                            error_msg += f"ERRORE nella creazione della maschera date per 'Data acq.': {str(e)}\n"
                            continue                        
                    else:
                        error_msg += "ERRORE nella valutazione della colonna 'Data acq.'\n"
                        continue             
                    
                    ### Maschera con date
                    mask_finale = mask_date_range & (mask_base | mask_legacy)
                    # mask_finale = (mask_base | mask_legacy)

                elif target_column == 'Visualizzati_in_OFA': # -> OK
                    mask_finale = mask_dict['DETAIL WORK ORDER']   
                
                elif target_column == 'Modificati_in_OFA': # -> OK
                    # Condizioni
                    mask_odm_upload_1 = mask_dict['WORKORDER HEADER UPDATE']
                    mask_odm_upload_2 = mask_dict['WORKORDER UPDATE RELEASE']
        
                    mask_finale = mask_odm_upload_1 | mask_odm_upload_2    
                
                elif target_column == 'Operazione_gestita_in_OFA': # -> OK
                    mask_finale = mask_dict['WORKORDER UPDATE OPERATION']                      

                elif target_column == 'Aggiunto_materiale': # -> OK
                    mask_finale = mask_dict['WORKORDER ADD COMPONENT'] 

                elif target_column == 'Gestione_materiale': # -> OK
                    # Condizioni
                    mask_odm_material_1 = mask_dict['WORKORDER DELETE COMPONENT']
                    mask_odm_material_2 = mask_dict['DETAIL WORK ORDER MATERIAL']
                    mask_odm_material_3 = mask_dict['WORKORDER UPDATE COMPONENT']
        
                    mask_finale = mask_odm_material_1 | mask_odm_material_2 | mask_odm_material_3    

                elif target_column == 'Allegati_in_OFA': # -> OK
                    # Condizioni
                    mask_odm_allegati_1 = mask_dict['OPEN TEXT DETAIL WORK ORDER OPERATION']
                    mask_odm_allegati_2 = mask_dict['OPEN TEXT UPLOAD ATTACHMEMT WORKORDER']
        
                    mask_finale = mask_odm_allegati_1 | mask_odm_allegati_2             

                elif target_column == 'Conferma_ore_in_OFA': # -> OK
                    # Condizioni
                    mask_odm_time_1 = mask_dict['CREATE TIME CONFERMATION']
                    mask_odm_time_2 = mask_dict['DETAIL WORK ORDER CONFERMATION']
        
                    mask_finale = mask_odm_time_1 | mask_odm_time_2 

                elif target_column == 'Chiusi_in_OFA':
                    # Condizioni
                    mask_chiusi = mask_dict['WORKORDER UPDATE TECO']                                     
                    # Verifica 'Stato sistema' in TECO oppure in CONC
                    mask_stato_sistema = df_result['Stato sistema'].str.contains('TECO|CONC', na=False, case=False)
                    
                    mask_finale = (mask_chiusi & mask_stato_sistema)

                else:
                    error_msg += f"ERRORE nella valutazione della colonna: {target_column} del df\n"
                    continue # Se rilevo un errore non procedo con l'applicazione della maschera ma continuo l'iterazione
                
                # Applica
                df_result.loc[mask_finale, target_column] = True
                logger.info(f"Colonna {target_column}: {df_result[target_column].sum()} True") # -> OK
            
            except Exception as e:
                logger.error(f"Errore nell'applicazione della maschera alla colonna: '{target_column}': {str(e)}")
                return False, None
        
        if error_msg:
                logger.error(f"Errori riscontrati: \n\t{error_msg}")
                return False, None

        # ========================================
        # Costruisce la colonne 'Numeratore'
        # ========================================  
        # Il numeratore assume valore 1 se almeno una delle colonne che descrivono gli stati OFA contiene un valore = 1
        # In pratica tutte le colonne sono in OR
        df_result['Numeratore'] = (
            df_result['Creati_in_OFA'] | 
            df_result['Visualizzati_in_OFA'] |
            df_result['Modificati_in_OFA'] |
            df_result['Operazione_gestita_in_OFA'] |                                
            df_result['Aggiunto_materiale'] |
            df_result['Gestione_materiale'] |
            df_result['Allegati_in_OFA'] |
            df_result['Conferma_ore_in_OFA'] |                                
            df_result['Chiusi_in_OFA']
        )  
        
        logger.info(f"Colonna {'Numeratore'}: {df_result['Numeratore'].sum()} True") # -> OK
        
        # ========================================
        # Costruisco la colonna 'Denominatore'
        # ========================================
        # Il denominatore assume valore 1 se il numeratore è 1 oppure se
        # la colonna 'Ordine' != "" E 
        # la colonna data modifica 'Mod. il' è vuota E
        # la data di inizio cardine è contenuta nel range di date selezionato nella GUI
        # altrimenti = 0
        
        # Costruisco la maschera a partire dai valori del numeratore
        mask_numeratore = df_result['Numeratore']

        # Definisco condizione per verificare che l'OdM abbia il campo data inizio cardine 'In. card.' entro il range di date selezionato nella GUI
        if 'In. card.' in df_result.columns:
            try:
                # Converti solo se non è già datetime
                if not pd.api.types.is_datetime64_any_dtype(df_result['In. card.']):
                    df_result['In. card.'] = pd.to_datetime(
                        df_result['In. card.'],
                        format='%d.%m.%Y',
                        errors='coerce'
                    )         
                # Maschera per il range di date (estremi compresi)
                mask_date_inizio_cardine = (
                    (df_result['In. card.'] >= pd_data_inizio) &
                    (df_result['In. card.'] <= pd_data_fine)
                )
            except Exception as e:
                error_msg += f"ERRORE nella creazione della maschera 'mask_date_inizio_cardine': {str(e)}\n"
                return False, None
        else:
            logger.error(f"ERRORE nella valutazione della colonna 'In. card.'\n")
            return False, None
        
        # Definisco condizione per verificare che l'OdM abbia il campo data modifica 'Data mod.' entro il range di date selezionato nella GUI
        if 'Data mod.' in df_result.columns:
            try:
                # Converti solo se non è già datetime
                if not pd.api.types.is_datetime64_any_dtype(df_result['Data mod.']):
                    df_result['Data mod.'] = pd.to_datetime(
                        df_result['Data mod.'],
                        format='%d.%m.%Y',
                        errors='coerce'
                    )             
                # Maschera per il range di date (estremi compresi)
                mask_date_modifica = (
                    (df_result['Data mod.'] >= pd_data_inizio) &
                    (df_result['Data mod.'] <= pd_data_fine)
                )
            except Exception as e:
                error_msg += f"ERRORE nella creazione della maschera 'mask_date_modifica': {str(e)}\n"
                return False, None
        else:
            logger.error(f"ERRORE nella valutazione della colonna 'Data mod.'\n")
            return False, None    
        
        # Condizione tipo OdM diverso da "M1"
        if 'Tp.' in df_result.columns:
            mask_tipo_ordine = (df_result['Tp.'] != 'M1') & (df_result['Tp.'].notna()) # Escludo i tipo 'M1' e i valori NaN
        else:
            logger.error(f"ERRORE nella valutazione della colonna 'Tp.'")
            return False, None

        # Condizione stato sistema OdM
        if 'Stato sistema' in df_result.columns:
            mask_stato_sistema = df_result['Stato sistema'].str.contains('TECO|CONC|RIL', na=False, case=False)       
        else:
            logger.error(f"ERRORE nella valutazione della colonna 'Stato sistema'")
            return False, None  

        df_result['Denominatore'] = False
        mask_finale = mask_numeratore | (mask_stato_sistema & mask_date_inizio_cardine & mask_date_modifica & mask_tipo_ordine)
        target_column = 'Denominatore'
        # Applica
        df_result.loc[mask_finale, target_column] = True
        logger.info(f"Colonna {target_column}: {df_result[target_column].sum()} True") # -> OK

        
        return True, df_result

    def df_add_columns(self, dict_df: dict[str, pd.DataFrame])  -> tuple[bool, dict[str, pd.DataFrame] | None]:
        """
        Aggiunge e modifica colonne nei dataframe secondo le regole definite

        Args:
            dict_df (dict[str, pd.DataFrame]): Dizionario contenente i dataframe da elaborare
        Returns:
            tuple[bool, dict[str, pd.DataFrame]]: Dizionario con i dataframe aggiornati
        """
        
        # Estrai i singoli dataframe per passarli alla funzione di elaborazione
        df_AFKO = dict_df["df_AFKO"]
        df_excel_normalized = dict_df["df_excel_normalized"] 
        df_IW29 = dict_df["df_IW29"]
        df_IW39 = dict_df["df_IW39"]
        df_plants = dict_df["df_plants"]

        ############################################
        # ----- Modifica df_excel_normalized ------
        ############################################

        print(f"\nInizio l'aggiornamento delle colonne nel df_excel_normalized...")
        print(f"\n\tInserimento colonna 'functionalLocation'...")
        try:
            # Inserisco una nuova colonna "FL_AdM" nel df_excel_normalized, ricavo la sede tecnica dal df_IW29
            df_excel_normalized["FL_AdM"] = df_excel_normalized["AdM"].map(df_IW29.set_index("Avviso")["Sede tecnica"])
            #self.log_manager.log(f"Sede tecnica AdM inserita correttamente dal df_IW29", "success")
            # Inserisco una nuova colonna "FL_OdM" nel df_excel_normalized, ricavo la sede tecnica dal df_IW39
            df_excel_normalized["FL_OdM"] = df_excel_normalized["OdM"].map(df_IW39.set_index("Ordine")["Sede tecnica"])
            #self.log_manager.log(f"Sede tecnica OdM inserita correttamente dal df_IW39", "success")
            # Verifico che non ci siano righe che contengono valori in entrambe le colonne FL_AdM e FL_OdM
            if (df_excel_normalized['FL_AdM'].notna() & df_excel_normalized['FL_OdM'].notna()).any():
                # Se esiste una riga che contiene entrambe i valori stampo un messaggio di errore e la riga
                print(f"Ci sono righe con valori in entrambe le colonne 'FL_AdM' e 'FL_OdM'!!!")
                print(df_excel_normalized[df_excel_normalized['FL_AdM'].notna() & df_excel_normalized['FL_OdM'].notna()])
                raise ValueError("Ci sono righe con valori in entrambe le colonne 'FL_AdM' e 'FL_OdM'")
            # Creo la colonna 'functionalLocation' combinando le due colonne FL_AdM e FL_OdM
            else:
                df_excel_normalized['functionalLocation'] = df_excel_normalized['FL_AdM'].fillna(df_excel_normalized['FL_OdM'])
                print(f"\n\tColonna 'functionalLocation' inserita correttamente...")
                #self.log_manager.log(f"Combinazione completata!", "success")
      
            ### Ricavo le colonne Country, Technology nel df df_excel_normalized a partire dal valore della colonna 'functionalLocation' ricavato nel punto precedente
            print(f"\n\tInserimento delle colonne 'Country' e 'Technology'...")
            # Crea colonna country (primi due caratteri)
            df_excel_normalized['country'] = df_excel_normalized['functionalLocation'].apply(
                lambda x: x[:2] if pd.notna(x) and isinstance(x, str) and len(x) >= 3 else None)
            #self.log_manager.log("Colonna 'country' inserita correttamente", "success")

            # Crea colonna tecnologia (primi due caratteri)
            df_excel_normalized['tecnology'] = df_excel_normalized['functionalLocation'].apply(
                lambda x: x[2] if pd.notna(x) and isinstance(x, str) and len(x) >= 3 else None)
            #self.log_manager.log("Colonna 'tecnology' inserita correttamente", "success")
            print(f"\n\tColonne 'country' e 'tecnology' inserite correttamente...")

            # Aggiungo la colonna 'tec_ext' nel df_excel_normalized
            print(f"\n\tInserimento della colonna 'tec_ext'...")
            def extract_and_transform_tech(functional_location):
                """Estrae il primo carattere e lo trasforma in tecnologia"""
                if pd.isna(functional_location) or not isinstance(functional_location, str) or len(functional_location) == 0:
                    return None
                
                first_char = functional_location[2].upper()
                
                mapping = {
                    'S': 'Solar',
                    'E': 'Bess',
                    'W': 'Wind'
                }
            
                return mapping.get(first_char, None)

            # Applica tutto insieme
            df_excel_normalized['tec_ext'] = df_excel_normalized['functionalLocation'].apply(extract_and_transform_tech)
            #self.log_manager.log("Colonna 'tec_ext' inserita correttamente", "success")
            print("Colonna 'tec_ext' inserita correttamente.")

        except Exception as e:
            print("Errore durante l'aggiornamento delle colonne nel df_excel_normalized:", str(e))
            #self.log_manager.log(f"Errore durante l'aggiornamento della FL in df_excel_normalizedil: {str(e)}", "error")
            return False, None 
        
        # # Salvo il df_excel_normalized aggiornato
        # success = Excel_file_mng.save_excel_file_advanced(df_excel_normalized, "df_excel_normalized_updated.xlsx")
        
        ############################################
        # ----- Modifica df_IW29 ------
        ############################################

        print(f"\nInizio l'aggiornamento delle colonne nel df_IW29...")
        try:
            # Aggiungo la colonna 'OdM_data_inizio_cardine' nel df_IW29 ricavandolo dal df_AFKO
            print(f"\n\tInserimento colonna 'OdM_data_inizio_cardine'...")
            # Crea dizionario di lookup per una singola colonna
            lookup_dict = df_AFKO.set_index('Ordine')['Data inizio cardine'].to_dict()
            df_IW29['OdM_data_inizio_cardine'] = df_IW29['Ordine'].map(lookup_dict)
            print(f"\n\tColonna 'OdM_data_inizio_cardine' inserita correttamente...")

            # Aggiungo la colonna 'Stato_sistema' nel df_IW29 ricavandolo dal df_IW39
            print(f"\n\tInserimento colonna 'OdM_stato_sistema'...")
            # Crea dizionario di lookup per una singola colonna
            lookup_dict = df_IW39.set_index('Ordine')['Stato sistema'].to_dict()
            df_IW29['OdM_stato_sistema'] = df_IW29['Ordine'].map(lookup_dict)
            print(f"\n\tColonna 'OdM_stato_sistema' inserita correttamente...")
            
            # Aggiungo la colonna 'PlantName' nel df_IW29 ricavandolo dal df_plants
            print(f"\n\tInserimento colonna 'Plant_name'...")
            # Crea dizionario di lookup per una singola colonna
            lookup_dict = df_plants.set_index('Functional Loc.')['Description'].to_dict()
            df_IW29['Plant_name'] = df_IW29['Sede tecnica'].str[:8].map(lookup_dict)
            print(f"\n\tColonna 'Plant_name' inserita correttamente...")

            # Aggiungo la colonna 'PlantName' nel df_IW29 ricavandolo dal df_plants
            print(f"\n\tInserimento colonna 'Strategy'...")
            # Crea dizionario di lookup per una singola colonna
            lookup_dict = df_plants.set_index('Functional Loc.')['Strategy Semplified'].to_dict()
            df_IW29['Strategy'] = df_IW29['Sede tecnica'].str[:8].map(lookup_dict)
            print(f"\n\tColonna 'Strategy' inserita correttamente...")

            # Aggiungo la colonna 'Tecnologia' nel df_IW29
            print(f"\n\tInserimento colonna 'Tecnologia'...")
            # Prelevo il terzo carattere della colonna 'Sede tecnica'
            df_IW29['Tecnologia'] = df_IW29['Sede tecnica'].str[2]  # Indice 2 = terzo carattere
            print(f"\n\tColonna 'Tecnologia' inserita correttamente...")

        except Exception as e:
            print("Errore durante l'aggiornamento delle colonne nel df_IW29:", str(e))
            #self.log_manager.log(f"Errore durante l'aggiornamento della FL in df_excel_normalizedil: {str(e)}", "error")
            return False, None             

        ############################################
        # ----- Modifica df_IW39 ------
        ############################################

        print(f"\nInizio l'aggiornamento delle colonne nel df_IW39...")
        try:
            print(f"\n\tInserimento colonna 'Strategy'...")
            # Aggiungo la colonna 'PlantName' nel df_IW39 ricavandolo dal df_plants
            # Crea dizionario di lookup per una singola colonna
            lookup_dict = df_plants.set_index('Functional Loc.')['Strategy Semplified'].to_dict()
            df_IW39['Strategy'] = df_IW39['Sede tecnica'].str[:8].map(lookup_dict)
            print(f"\n\tColonna 'Strategy' inserita correttamente...")

            # Aggiungo la colonna 'Tecnologia' nel df_IW39
            print(f"\n\tInserimento colonna 'Tecnologia'...")
            # Prelevo il terzo carattere della colonna 'Sede tecnica'
            df_IW39['Tecnologia'] = df_IW39['Sede tecnica'].str[2]  # Indice 2 = terzo carattere   
            print(f"\n\tColonna 'Tecnologia' inserita correttamente...")

            # Aggiungo la colonna 'PlantName' nel df_IW39 ricavandolo dal df_plants
            print(f"\n\tInserimento colonna 'Plant_name'...")
            # Crea dizionario di lookup per una singola colonna
            lookup_dict = df_plants.set_index('Functional Loc.')['Description'].to_dict()
            df_IW39['Plant_name'] = df_IW39['Sede tecnica'].str[:8].map(lookup_dict)
            print(f"\n\tColonna 'Plant_name' inserita correttamente...")

        except Exception as e:
            print("Errore durante l'aggiornamento delle colonne nel df_IW39:", str(e))
            #self.log_manager.log(f"Errore durante l'aggiornamento della FL in df_excel_normalizedil: {str(e)}", "error")
            return False, None 

        # Aggiorno il dizionario dei dataframe
        dict_df = {}
        dict_df["df_excel_normalized"] = df_excel_normalized
        dict_df["df_IW29"] = df_IW29  
        dict_df["df_IW39"] = df_IW39

        print(f"\nAggiornamento delle colonne completato con successo...")
        return True, dict_df

    def process_dataframes(self, dict_df: dict[str, pd.DataFrame]) -> tuple[bool, dict[str, pd.DataFrame] | None]:
        """
        Esegue l'elaborazione dei dataframe secondo le regole definite

        Args:
            dict_df (dict[str, pd.DataFrame]): Dizionario contenente i dataframe da elaborare
        Returns:
            bool: True se l'elaborazione è andata a buon fine, False altrimenti
        """         
        # Estrai i singoli dataframe per passarli alla funzione di elaborazione
        df_excel_normalized = dict_df["df_excel_normalized"]
        df_IW29 = dict_df["df_IW29"]
        df_IW39 = dict_df["df_IW39"]

        # ----- Avvio l'elaborazione degli avvisi -----
        result, df_AdM = self.process_iw29_with_ofa_actions(df_IW29, df_excel_normalized)
        if result == False:
            logger.info(f"Errore durante l'elaborazione del df IW29") # -> OK
            return False, None
        
        # # Salvo il df per utilizzarlo nella PowerBI
        # success = Excel_file_mng.save_excel_file_advanced(df_AdM, "AdM_PowerBI.xlsx")

        # Filtro in base al valore contenuto nella colonna 'Strategy' prima di costruire la pivot
        df_filtrato = df_AdM[df_AdM['Strategy'].isin(['IH', 'SM', '']) & df_AdM['Strategy'].notna()]

        # Multiple funzioni di aggregazione
        pivot_multi = df_filtrato.pivot_table(
            index=['Tecnologia', 'Pse'],  # Righe gerarchiche
            values=['Numeratore', 'Denominatore'],  # Colonna da sommare
            aggfunc='sum'  # Funzione di aggregazione
        ).assign(KPI=lambda x: x['Numeratore'] / x['Denominatore'] * 100)

        print(f"\nPivot con multiple aggregazioni:")
        print(pivot_multi.to_string())

        # ----- Avvio l'elaborazione degli ordini -----
        result, df_OdM = self.process_iw39_with_ofa_actions(df_IW39, df_excel_normalized)
        if result == False:
            logger.info(f"Errore durante l'elaborazione del df IW39") # -> OK
            return False, None

        # # Salvo il df per utilizzarlo nella PowerBI
        # success = Excel_file_mng.save_excel_file_advanced(df_OdM, "OdM_PowerBI.xlsx")

        # Filtro in base al valore contenuto nella colonna 'Strategy' prima di costruire la pivot
        df_filtrato = df_OdM[df_OdM['Strategy'].isin(['IH', 'SM', '']) & df_OdM['Strategy'].notna()]

        # Multiple funzioni di aggregazione
        pivot_multi = df_filtrato.pivot_table(
            index=['Tecnologia', 'Pse'],  # Righe gerarchiche
            values=['Numeratore', 'Denominatore'],  # Colonna da sommare
            aggfunc='sum'  # Funzione di aggregazione
        ).assign(KPI=lambda x: x['Numeratore'] / x['Denominatore'] * 100)

        print(f"\nPivot con multiple aggregazioni:")
        print(pivot_multi.to_string())  
        
        # Aggiorno il dizionario dei dataframe
        dict_df_elaborated = {}
        dict_df_elaborated["df_AdM"] = df_AdM
        dict_df_elaborated["df_OdM"] = df_OdM

        return True, dict_df_elaborated
