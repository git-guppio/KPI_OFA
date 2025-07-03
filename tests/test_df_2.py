import pandas as pd
import numpy as np
import logging

# Configurazione logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

def process_iw29_with_ofa_actions(df_IW29, df_excel_normalized, df_IW39):
    """
    Processa il DataFrame df_IW29 aggiungendo colonne OFA basate sui dati di df_excel_normalized
    
    Args:
        df_IW29 (pd.DataFrame): DataFrame IW29 con colonne 'Avviso' e 'Ordine'
        df_excel_normalized (pd.DataFrame): DataFrame normalizzato con colonne 'AdM' e 'action'
        df_IW39 (pd.DataFrame): DataFrame IW39 con colonne 'Ordine' e 'Stato sistema'
    
    Returns:
        pd.DataFrame: DataFrame df_IW29 aggiornato con le nuove colonne OFA
        
    ELABORAZIONE NOTIFICATION:
    - CREATE NOTIFICATION -> Creati_in_OFA
    - DETAIL NOTIFICATION -> Visualizzati_in_OFA
    - UPDATE NOTIFICATION -> Modificati_in_OFA
    - OPEN TEXT UPLOAD ATTACHMEMT NOTIFICATION/OPEN TEXT UPLOAD ATTACHMEMT -> Allegati_in_OFA
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
    
    # Verifica df_IW39 per gestione WORKORDER UPDATE TECO
    if df_IW39 is not None and not df_IW39.empty:
        if 'Ordine' not in df_IW39.columns or 'Stato sistema' not in df_IW39.columns:
            logger.warning("Colonne 'Ordine' o 'Stato sistema' non trovate in df_IW39 - WORKORDER UPDATE TECO non sarà processato")
            df_IW39 = None
    else:
        logger.warning("df_IW39 è vuoto o None - WORKORDER UPDATE TECO non sarà processato")
    
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
        'CLOSED NOTIFICATION': 'Chiusi_in_OFA'
    }
    
    # Crea un dizionario di lookup per velocizzare le ricerche
    # Converte AdM prima in intero poi in stringa per garantire coerenza
    lookup_dict = {}
    for idx, row in df_excel_normalized.iterrows():
        try:
            # Converte prima in intero per rimuovere decimali, poi in stringa
            if pd.isna(row['AdM']):
                continue  # Salta valori NaN
            
            adm_int = int(float(row['AdM']))  # float -> int per gestire "1001.0"
            adm_value = str(adm_int)  # int -> string risulta in "1001"
            action_value = str(row['action']).strip()
            lookup_dict[adm_value] = action_value
            
        except (ValueError, TypeError):
            logger.warning(f"Valore AdM non convertibile: {row['AdM']} alla riga {idx}")
            continue
    
    logger.info(f"Creato dizionario di lookup con {len(lookup_dict)} elementi")
    
    # Trova avvisi con WORKORDER UPDATE TECO e crea lookup per ordini
    workorder_teco_avvisi = [adm for adm, action in lookup_dict.items() if action == 'WORKORDER UPDATE TECO']
    logger.info(f"Trovati {len(workorder_teco_avvisi)} avvisi con WORKORDER UPDATE TECO: {workorder_teco_avvisi}")
    
    # Crea dizionario di lookup per ordini da df_IW39 (se disponibile)
    ordine_stato_lookup = {}
    if df_IW39 is not None:
        for idx, row in df_IW39.iterrows():
            try:
                if pd.isna(row['Ordine']):
                    continue
                
                ordine_int = int(float(row['Ordine']))  # Converti come per AdM
                ordine_str = str(ordine_int)
                stato_sistema = str(row['Stato sistema']).strip().upper()  # Normalizza maiuscole
                ordine_stato_lookup[ordine_str] = stato_sistema
                
            except (ValueError, TypeError):
                logger.warning(f"Valore Ordine non convertibile in df_IW39: {row['Ordine']} alla riga {idx}")
                continue
        
        logger.info(f"Creato lookup ordini con {len(ordine_stato_lookup)} elementi")
    
    # Contatori per statistiche
    stats = {
        'totale_righe': len(df_result),
        'avvisi_validi': 0,
        'avvisi_non_numerici': 0,
        'avvisi_non_trovati': 0,
        'workorder_teco_processati': 0,
        'workorder_teco_chiusi': 0,
        'azioni_processate': {col: 0 for col in ofa_columns}
    }
    
    # Processa ogni riga del DataFrame
    for idx, row in df_result.iterrows():
        avviso_value = row['Avviso']
        
        # Verifica che il valore sia un numero intero
        try:
            # Converte in stringa e rimuove spazi
            if pd.isna(avviso_value):
                stats['avvisi_non_numerici'] += 1
                continue
                
            # Prova a convertire in intero
            if isinstance(avviso_value, (int, float)):
                avviso_int = int(avviso_value)
                avviso_str = str(avviso_int)
            else:
                avviso_str = str(avviso_value).strip()
                avviso_int = int(float(avviso_str))  # Gestisce "123.0" -> 123
                avviso_str = str(avviso_int)
            
            stats['avvisi_validi'] += 1
            
        except (ValueError, TypeError):
            logger.warning(f"Riga {idx}: Valore 'Avviso' non numerico: {avviso_value}")
            stats['avvisi_non_numerici'] += 1
            continue
        
        # Cerca il valore nel dizionario di lookup
        if avviso_str in lookup_dict:
            action = lookup_dict[avviso_str]
            
            # Gestione WORKORDER UPDATE TECO (caso speciale)
            if action == 'WORKORDER UPDATE TECO':
                stats['workorder_teco_processati'] += 1
                
                # Recupera il valore dell'ordine per questo avviso
                try:
                    ordine_value = row['Ordine']
                    if pd.isna(ordine_value):
                        logger.warning(f"Riga {idx}: Ordine nullo per avviso {avviso_str} con WORKORDER UPDATE TECO")
                        continue
                    
                    # Converte ordine come fatto per avviso
                    if isinstance(ordine_value, (int, float)):
                        ordine_int = int(ordine_value)
                        ordine_str = str(ordine_int)
                    else:
                        ordine_str = str(ordine_value).strip()
                        ordine_int = int(float(ordine_str))
                        ordine_str = str(ordine_int)
                    
                    # Cerca lo stato sistema nell'IW39
                    if ordine_str in ordine_stato_lookup:
                        stato_sistema = ordine_stato_lookup[ordine_str]
                        
                        # Verifica se contiene TECO o CONC
                        if 'TECO' in stato_sistema or 'CONC' in stato_sistema:
                            df_result.at[idx, 'Chiusi_in_OFA'] = True
                            stats['azioni_processate']['Chiusi_in_OFA'] += 1
                            stats['workorder_teco_chiusi'] += 1
                            logger.debug(f"Riga {idx}: Avviso {avviso_str} -> Ordine {ordine_str} -> Stato '{stato_sistema}' -> Chiusi_in_OFA")
                        else:
                            logger.debug(f"Riga {idx}: Ordine {ordine_str} con stato '{stato_sistema}' non contiene TECO/CONC")
                    else:
                        logger.debug(f"Riga {idx}: Ordine {ordine_str} non trovato in df_IW39")
                        
                except (ValueError, TypeError):
                    logger.warning(f"Riga {idx}: Valore 'Ordine' non convertibile: {row['Ordine']}")
                    continue
            
            # Gestione azioni standard
            elif action in action_mapping:
                target_column = action_mapping[action]
                df_result.at[idx, target_column] = True
                stats['azioni_processate'][target_column] += 1
                
                logger.debug(f"Riga {idx}: Avviso {avviso_str} -> Azione '{action}' -> {target_column}")
            else:
                logger.warning(f"Riga {idx}: Azione non riconosciuta: '{action}'")
        else:
            logger.debug(f"Riga {idx}: Avviso {avviso_str} non trovato in df_excel_normalized")
            stats['avvisi_non_trovati'] += 1
    
    # Stampa statistiche
    print_processing_stats(stats)
    
    return df_result


def print_processing_stats(stats):
    """
    Stampa le statistiche del processamento
    
    Args:
        stats (dict): Dizionario con le statistiche
    """
    print("\n" + "="*60)
    print("STATISTICHE PROCESSAMENTO df_IW29")
    print("="*60)
    
    print(f"Totale righe processate: {stats['totale_righe']}")
    print(f"Avvisi validi (numerici): {stats['avvisi_validi']}")
    print(f"Avvisi non numerici: {stats['avvisi_non_numerici']}")
    print(f"Avvisi non trovati in lookup: {stats['avvisi_non_trovati']}")
    print(f"WORKORDER UPDATE TECO processati: {stats['workorder_teco_processati']}")
    print(f"WORKORDER UPDATE TECO -> Chiusi: {stats['workorder_teco_chiusi']}")
    
    print("\nAzioni processate per colonna:")
    for col, count in stats['azioni_processate'].items():
        print(f"  {col}: {count}")
    
    print("="*60)


def analyze_dataframes_before_processing(df_IW29, df_excel_normalized):
    """
    Analizza i DataFrame prima del processamento per identificare potenziali problemi
    
    Args:
        df_IW29 (pd.DataFrame): DataFrame IW29
        df_excel_normalized (pd.DataFrame): DataFrame normalizzato
    """
    print("\n" + "="*60)
    print("ANALISI PRE-PROCESSAMENTO")
    print("="*60)
    
    # Analisi df_IW29
    print(f"\ndf_IW29:")
    print(f"  Shape: {df_IW29.shape}")
    print(f"  Colonne: {list(df_IW29.columns)}")
    
    if 'Avviso' in df_IW29.columns:
        avviso_col = df_IW29['Avviso']
        print(f"  Colonna 'Avviso':")
        print(f"    Tipo dati: {avviso_col.dtype}")
        print(f"    Valori nulli: {avviso_col.isnull().sum()}")
        print(f"    Valori unici: {avviso_col.nunique()}")
        print(f"    Esempi: {list(avviso_col.head().values)}")
    
    # Analisi df_excel_normalized
    print(f"\ndf_excel_normalized:")
    print(f"  Shape: {df_excel_normalized.shape}")
    print(f"  Colonne: {list(df_excel_normalized.columns)}")
    
    if 'AdM' in df_excel_normalized.columns and 'action' in df_excel_normalized.columns:
        adm_col = df_excel_normalized['AdM']
        action_col = df_excel_normalized['action']
        
        print(f"  Colonna 'AdM':")
        print(f"    Tipo dati: {adm_col.dtype}")
        print(f"    Valori nulli: {adm_col.isnull().sum()}")
        print(f"    Valori unici: {adm_col.nunique()}")
        print(f"    Esempi: {list(adm_col.head().values)}")
        
        print(f"  Colonna 'action':")
        print(f"    Tipo dati: {action_col.dtype}")
        print(f"    Valori nulli: {action_col.isnull().sum()}")
        print(f"    Azioni uniche: {action_col.nunique()}")
        print(f"    Azioni disponibili:")
        for action in sorted(action_col.unique()):
            count = (action_col == action).sum()
            print(f"      '{action}': {count}")


def esempio_utilizzo_completo(df_IW29, df_excel_normalized, df_IW39=None):
    """
    Esempio completo di utilizzo delle funzioni
    
    Args:
        df_IW29 (pd.DataFrame): DataFrame IW29
        df_excel_normalized (pd.DataFrame): DataFrame normalizzato
        df_IW39 (pd.DataFrame): DataFrame IW39 (opzionale)
    
    Returns:
        pd.DataFrame: DataFrame processato
    """
    print("INIZIO PROCESSAMENTO df_IW29 CON AZIONI OFA")
    print("="*60)
    
    # Analisi preliminare
    analyze_dataframes_before_processing(df_IW29, df_excel_normalized)
    
    # Processamento
    df_result = process_iw29_with_ofa_actions(df_IW29, df_excel_normalized, df_IW39)
    
    # Verifica risultati
    if not df_result.empty:
        print(f"\nDataFrame risultante:")
        print(f"  Shape: {df_result.shape}")
        print(f"  Nuove colonne aggiunte:")
        ofa_columns = ['Creati_in_OFA', 'Visualizzati_in_OFA', 'Modificati_in_OFA', 
                       'Allegati_in_OFA', 'Chiusi_in_OFA']
        
        for col in ofa_columns:
            if col in df_result.columns:
                true_count = df_result[col].sum()
                print(f"    {col}: {true_count} True su {len(df_result)} righe")
        
        # Mostra alcune righe di esempio
        print(f"\nPrime 5 righe con le nuove colonne:")
        cols_to_show = ['Avviso', 'Ordine'] + ofa_columns
        available_cols = [col for col in cols_to_show if col in df_result.columns]
        print(df_result[available_cols].head().to_string())
    
    return df_result


# Funzione di test per verificare la logica
def test_function_with_sample_data():
    """
    Funzione di test con dati campione per verificare la logica
    """
    print("\n" + "="*60)
    print("TEST CON DATI CAMPIONE")
    print("="*60)
    
    # Crea dati di test che includono WORKORDER UPDATE TECO
    df_IW29_test = pd.DataFrame({
        'Avviso': [1001, 1002, 1003, '1004', 1005.0, 'invalid', None, 1006, 1007],
        'Ordine': [2001, 2002, 2003, 2004, 2005, 2006, 2007, 2008, 2009],
        'Altri_Dati': ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I']
    })
    
    # DataFrame con valori float che causavano il problema
    df_excel_normalized_test = pd.DataFrame({
        'AdM': [1001.0, 1002.0, 1003.0, 1004.0, 1005.0, 1006.0, 1007.0],  # Valori float
        'action': ['CREATE NOTIFICATION', 'DETAIL NOTIFICATION', 'UPDATE NOTIFICATION',
                  'OPEN TEXT UPLOAD ATTACHMEMT', 'CLOSED NOTIFICATION', 'UNKNOWN ACTION', 'WORKORDER UPDATE TECO']
    })
    
    # DataFrame IW39 per testare WORKORDER UPDATE TECO
    df_IW39_test = pd.DataFrame({
        'Ordine': [2001.0, 2002.0, 2003.0, 2007.0, 2008.0, 2009.0],
        'Stato sistema': ['APERTO', 'IN LAVORO', 'TECO COMPLETATO', 'CONC CONFERMATO', 'IN PROGRESS', 'TECO FINALE']
    })
    
    print("Dati di test creati:")
    print("df_IW29_test:")
    print(df_IW29_test.to_string())
    print("\ndf_excel_normalized_test (nota i valori float in AdM):")
    print(df_excel_normalized_test.to_string())
    print(f"Tipo dati AdM: {df_excel_normalized_test['AdM'].dtype}")
    print("\ndf_IW39_test:")
    print(df_IW39_test.to_string())
    
    # Dimostra il problema e la soluzione
    print(f"\nDIMOSTRAZIONE DEL PROBLEMA RISOLTO:")
    print(f"Valore AdM originale: {df_excel_normalized_test['AdM'].iloc[0]} (tipo: {type(df_excel_normalized_test['AdM'].iloc[0])})")
    
    # Conversione corretta
    adm_val = df_excel_normalized_test['AdM'].iloc[0]
    adm_int = int(float(adm_val))
    adm_str = str(adm_int)
    print(f"Dopo conversione int->str: '{adm_str}'")
    
    # Conversione avviso
    avviso_val = df_IW29_test['Avviso'].iloc[0] 
    avviso_int = int(avviso_val)
    avviso_str = str(avviso_int)
    print(f"Avviso dopo conversione: '{avviso_str}'")
    print(f"Match trovato: {adm_str == avviso_str}")
    
    # Test WORKORDER UPDATE TECO
    print(f"\nTEST WORKORDER UPDATE TECO:")
    print(f"Avviso 1007 dovrebbe essere mappato a WORKORDER UPDATE TECO")
    print(f"Ordine 2009 ha stato 'TECO FINALE' -> dovrebbe settare Chiusi_in_OFA = True")
    
    # Esegui il test
    result = esempio_utilizzo_completo(df_IW29_test, df_excel_normalized_test, df_IW39_test)
    
    return result


if __name__ == "__main__":
    # Esegui il test
    test_result = test_function_with_sample_data()