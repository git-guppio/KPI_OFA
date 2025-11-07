import pandas as pd
import sys
import os

# Aggiungi la directory principale al PYTHONPATH
current_dir = os.path.dirname(os.path.abspath(__file__))  # tests directory
project_root = os.path.dirname(current_dir)  # directory principale
sys.path.insert(0, project_root)

import kpi_ofa.config.constants as constants
from kpi_ofa.core.utilis import DataFrameTools as DF_Tools

# def invert_field_names_map(transaction_code):
#     """
#     Inverte il dizionario per una specifica transazione.
#     Da field_name -> [varianti] a variante -> field_name
    
#     Args:
#         transaction_code: codice transazione (es. "IW29", "IW39")
        
#     Returns:
#         dict: dizionario invertito {variante: field_name}
#     """
#     inverted_map = {}
#     field_names_map = constants.SAP_tech_field_names_map

#     if field_names_map is None:
#         print("Attenzione: dizionario dei field names non definito")
#         return inverted_map

#     if transaction_code not in field_names_map:
#         print(f"Attenzione: transazione {transaction_code} non trovata nel dizionario")
#         return inverted_map
    
#     for field_name, variants in field_names_map[transaction_code].items():
#         for variant in variants:
#             inverted_map[variant] = field_name
    
#     return inverted_map


# def normalize_column_names(df, transaction_code, use_friendly_names=False):
#     """
#     Normalizza i nomi delle colonne di un dataframe usando i field names standard.
    
#     Args:
#         df: DataFrame pandas da normalizzare
#         transaction_code: codice transazione (es. "IW29", "IW39")
#         use_friendly_names: se True, usa nomi user-friendly invece dei field names tecnici
        
#     Returns:
#         DataFrame con colonne rinominate
#     """
#     if df is None or df.empty:
#         return df
    
#     SAP_user_field_names_map = constants.SAP_user_field_names_map

#     # Ottieni il dizionario invertito per questa transazione
#     inverted_map = invert_field_names_map(transaction_code)
    
#     # Crea il dizionario di rinominazione per le colonne presenti nel dataframe
#     rename_dict = {}
#     columns_not_found = []
    
#     for col in df.columns:
#         if col in inverted_map:
#             field_name = inverted_map[col]
            
#             # Se richiesto, usa i nomi user-friendly
#             if use_friendly_names and transaction_code in SAP_user_field_names_map:
#                 if field_name in SAP_user_field_names_map[transaction_code]:
#                     rename_dict[col] = SAP_user_field_names_map[transaction_code][field_name]
#                 else:
#                     rename_dict[col] = field_name  # Fallback al field name tecnico
#             else:
#                 rename_dict[col] = field_name
#         else:
#             columns_not_found.append(col)
    
#     # Avvisa se ci sono colonne non mappate
#     if columns_not_found:
#         print(f"Attenzione: colonne non trovate nel mapping per {transaction_code}: {columns_not_found}")
    
#     # Rinomina le colonne
#     df_normalized = df.rename(columns=rename_dict)
    
#     return df_normalized


# def verify_columns_compatibility(dataframes_list):
#     """
#     Verifica che tutti i dataframe abbiano le stesse colonne.
    
#     Args:
#         dataframes_list: lista di DataFrame da verificare
        
#     Returns:
#         tuple: (bool: sono compatibili, set: colonne comuni, dict: differenze per df)
#     """
#     if not dataframes_list or len(dataframes_list) == 0:
#         return True, set(), {}
    
#     # Filtra dataframe None o vuoti
#     valid_dfs = [df for df in dataframes_list if df is not None and not df.empty]
    
#     if len(valid_dfs) == 0:
#         return True, set(), {}
    
#     # Ottieni le colonne del primo dataframe come riferimento
#     reference_columns = set(valid_dfs[0].columns)
    
#     # Verifica compatibilità
#     are_compatible = True
#     differences = {}
    
#     for idx, df in enumerate(valid_dfs):
#         df_columns = set(df.columns)
        
#         if df_columns != reference_columns:
#             are_compatible = False
#             missing = reference_columns - df_columns
#             extra = df_columns - reference_columns
            
#             differences[f"DataFrame_{idx}"] = {
#                 "colonne_mancanti": list(missing),
#                 "colonne_extra": list(extra)
#             }
    
#     return are_compatible, reference_columns, differences


# def normalize_and_concat(dataframes_dict, transaction_code, use_friendly_names=False, strict_mode=True):
#     """
#     Normalizza e concatena una lista di dataframe con verifica di compatibilità.
    
#     Args:
#         dataframes_dict: dizionario o lista di dataframe (es. iw29.values())
#         transaction_code: codice transazione per il mapping
#         use_friendly_names: se True, usa nomi user-friendly
#         strict_mode: se True, blocca la concatenazione se le colonne non sono compatibili
        
#     Returns:
#         DataFrame concatenato con colonne normalizzate, o None se fallisce
#     """
#     if not dataframes_dict:
#         print("Nessun dataframe da concatenare")
#         return None
    
#     # Normalizza tutti i dataframe
#     normalized_dfs = []
#     for idx, df in enumerate(dataframes_dict):
#         if df is not None and not df.empty:
#             normalized_df = normalize_column_names(df, transaction_code, use_friendly_names)
#             normalized_dfs.append(normalized_df)
    
#     if not normalized_dfs:
#         print("Nessun dataframe valido dopo la normalizzazione")
#         return None
    
#     # Verifica compatibilità colonne
#     are_compatible, common_columns, differences = verify_columns_compatibility(normalized_dfs)
    
#     print(f"\n{'='*60}")
#     print(f"VERIFICA COMPATIBILITÀ - Transazione: {transaction_code}")
#     print(f"{'='*60}")
#     print(f"Numero dataframe da concatenare: {len(normalized_dfs)}")
#     print(f"Colonne comuni: {sorted(common_columns)}")
    
#     if are_compatible:
#         print("✓ Tutti i dataframe hanno le stesse colonne")
        
#         # Concatena
#         result_df = pd.concat(normalized_dfs, ignore_index=True)
#         print(f"✓ Concatenazione completata: {len(result_df)} righe totali")
#         print(f"{'='*60}\n")
#         return result_df
#     else:
#         print("✗ ATTENZIONE: I dataframe hanno colonne diverse!")
#         print("\nDifferenze rilevate:")
#         for df_name, diff in differences.items():
#             print(f"\n  {df_name}:")
#             if diff["colonne_mancanti"]:
#                 print(f"    Colonne mancanti: {diff['colonne_mancanti']}")
#             if diff["colonne_extra"]:
#                 print(f"    Colonne extra: {diff['colonne_extra']}")
        
#         if strict_mode:
#             print(f"\n✗ Concatenazione BLOCCATA (strict_mode=True)")
#             print(f"{'='*60}\n")
#             return None
#         else:
#             print(f"\n⚠ Concatenazione PERMESSA (strict_mode=False)")
#             print("  Nota: pandas riempirà con NaN le colonne mancanti")
#             result_df = pd.concat(normalized_dfs, ignore_index=True)
#             print(f"✓ Concatenate {len(result_df)} righe totali")
#             print(f"{'='*60}\n")
#             return result_df



if __name__ == "__main__":
    
    df_utils = DF_Tools()

# Creo dei df di test
## Crea il DataFrame
    df_1 = pd.DataFrame({
        'Avviso': [1, 2],
        'Mod. il': ['', ''],
        'Data cr.': ['31.10.2025', '13.10.2025'],
        'Descrizione': ['813 MXW-MXA2-X1-02', '222 MXW-MXA3-X1-01'],
        'Tp.': ['Z2', 'Z2'],
        'Sede tecnica': ['MXW-MXA2-X1-02-PS', 'MXW-MXA3-X1-01-HG-ES'],
        'St.sist.': ['MAPE', 'MAPE'],
        'Ordine': ['', ''],
        'Pse': ['MX', 'MX'],
        'Sis.Legacy': ['POM', 'POM']
    })
    df_2 = pd.DataFrame({
        'Avviso': [3, 4],
        'Mod. il': ['', ''],
        'Data cr.': ['31.10.2025', '13.10.2025'],
        'Descrizione': ['813 MXW-MXA2-X1-02', '222 MXW-MXA3-X1-01'],
        'Tp.': ['Z2', 'Z2'],
        'Sede tecnica': ['MXW-MXA2-X1-02-PS', 'MXW-MXA3-X1-01-HG-ES'],
        'Stato sistema': ['MAPE', 'MAPE'],
        'Ordine': ['', ''],
        'Pse': ['MX', 'MX'],
        'Sis.Legacy': ['POM', 'POM']
    })
    df_3 = pd.DataFrame({
        'Avviso': [5, 6],
        'Mod.': ['', ''],
        'Data cr.': ['31.10.2025', '13.10.2025'],
        'Descrizione': ['813 MXW-MXA2-X1-02', '222 MXW-MXA3-X1-01'],
        'Tp.': ['Z2', 'Z2'],
        'Sede tecnica': ['MXW-MXA2-X1-02-PS', 'MXW-MXA3-X1-01-HG-ES'],
        'St.sist.': ['MAPE', 'MAPE'],
        'Ordine': ['', ''],
        'Pse': ['MX', 'MX'],
        'Sis.Legacy': ['POM', 'POM']
    })

# creo n dizionario di df
    diz_df = {"df_1": df_1, "df_2": df_2, "df_3": df_3}
    
    # Esegui il test

    # Esempio 1: Concatena con field names tecnici (QMNUM, AEDAT, ecc.)
    result_iw29_technical = df_utils.normalize_and_concat(diz_df.values(), "IW29", use_friendly_names=False)
    if result_iw29_technical is None:
        print("DataFrame non creato correttamente")
        sys.exit(1)
    print(result_iw29_technical)

    # Esempio 2: Concatena con nomi user-friendly (Avviso, Mod. il, ecc.)
    result_iw29_friendly = df_utils.normalize_and_concat(diz_df.values(), "IW29", use_friendly_names=True)
    if result_iw29_friendly is None:
        print("DataFrame non creato correttamente")
        sys.exit(1)
    print(result_iw29_friendly)




