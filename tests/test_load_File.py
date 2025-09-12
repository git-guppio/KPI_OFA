import sys
import os
import pandas as pd

# Aggiungi la root del progetto al path
project_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, project_root)

from tests.class_excel_file_mng import ExcelFileMng
from tests.class_df_processor import DfProcessor

import logging

# Configurazione logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)


def main():
    #Creo una istanza delle classi necessarie
    Excel_file_mng = ExcelFileMng(base_path="C:/Users/a259046/OneDrive - Enel Spa/SCRIPT AHK e VBA/GITHUB/KPI_OFA/data/test")
    # carico i file nel df
    if not Excel_file_mng.validate_path():
        print("Il percorso non è valido.")
        return
    result, df_dict = Excel_file_mng.load_all_files()
    if result:
        # creo una istanza di DfProcessor per elaborare i DataFrame
        df_processor = DfProcessor(df_dict)
        df_processor.get_dataframe_info()
    else:
        print("Errore nel caricamento dei file.")
        return
    
    # Converti tutti i campi dei df in stringhe
    print("\nInizio conversione dei DataFrame...")
    result, df_dict = df_processor.converti_tutti_df(df_dict)


    # Estrai i singoli dataframe per passarli alla funzione di elaborazione
    df_AFKO = df_dict[df_AFKO]
    df_excel_normalized = df_dict[df_excel_normalized]
    df_IW29 = df_dict[df_IW29]
    df_IW39 = df_dict[df_IW39]
    df_plants = df_dict[df_plants]

    # ----- Modifica df_excel_normalized ------
    try:
        # Inserisco una nuova colonna "FL_AdM" nel df_excel_normalized, ricavo la sede tecnica dal df_IW29
        df_excel_normalized["FL_AdM"] = df_excel_normalized["AdM"].map(df_IW29.set_index("Avviso")["Sede tecnica"])
        #self.log_manager.log(f"Sede tecnica AdM inserita correttamente dal df_IW29", "success")
        # Inserisco una nuova colonna "FL_OdM" nel df_excel_normalized, ricavo la sede tecnica dal df_IW39
        df_excel_normalized["FL_OdM"] = df_excel_normalized["OdM"].map(df_IW29.set_index("Avviso")["Sede tecnica"])
        #self.log_manager.log(f"Sede tecnica OdM inserita correttamente dal df_IW39", "success")
        # Verifico che non ci siano righe che contengono valori in entrambe le colonne FL_AdM e FL_OdM
        if (df_excel_normalized['FL_AdM'].notna() & df_excel_normalized['FL_OdM'].notna()).any():
            #self.log_manager.log(f"Ci sono righe con valori in entrambe le colonne", "error")
            print(df_excel_normalized[df_excel_normalized['FL_AdM'].notna() & df_excel_normalized['FL_OdM'].notna()])
        else:
            df_excel_normalized['functionalLocation'] = df_excel_normalized['FL_AdM'].fillna(df_excel_normalized['FL_OdM'])
            print(f"Combinazione completata!")
            #self.log_manager.log(f"Combinazione completata!", "success")
    except Exception as e:
        print(f"Errore durante l'aggiornamento della FL in df_excel_normalized: {str(e)}")
        #self.log_manager.log(f"Errore durante l'aggiornamento della FL in df_excel_normalized: {str(e)}", "error")
        return False, None            

    ### Ricavo le colonne Country, Tecnology nel df df_excel_normalized a partire dal valore della colonna 'functionalLocation' ricavato nel punto precedente
    print(f"\nInizio l'aggiornamento delle colonne Country e Tecnology nel df_excel_normalized...")
    #self.log_manager.log("Inserisco Country e Tecnology nel df_excel_normalized", "info")
    try:

        # Crea colonna country (primi due caratteri)
        df_excel_normalized['country'] = df_excel_normalized['functionalLocation'].apply(
            lambda x: x[:2] if pd.notna(x) and isinstance(x, str) and len(x) >= 3 else None)
        #self.log_manager.log("Colonna 'country' inserita correttamente", "success")

        # Crea colonna tecnologia (primi due caratteri)
        df_excel_normalized['tecnology'] = df_excel_normalized['functionalLocation'].apply(
            lambda x: x[3] if pd.notna(x) and isinstance(x, str) and len(x) >= 3 else None)
        #self.log_manager.log("Colonna 'tecnology' inserita correttamente", "success")
        print("Colonne 'country' e 'tecnology' inserite correttamente nel df_excel_normalized")

        # Crea colonna TL (terzo carattere)
        def extract_and_transform_tech(functional_location):
            """Estrae il primo carattere e lo trasforma in tecnologia"""
            if pd.isna(functional_location) or not isinstance(functional_location, str) or len(functional_location) == 0:
                return None
            
            first_char = functional_location[3].upper()
            
            mapping = {
                'S': 'Solar',
                'E': 'Bess',
                'W': 'Wind'
            }
        
            return mapping.get(first_char, None)

        # Applica tutto insieme
        df_excel_normalized['tec_ext'] = df_excel_normalized['functionalLocation'].apply(extract_and_transform_tech)
        #self.log_manager.log("Colonna 'tec_ext' inserita correttamente", "success")
        print("Colonna 'tec_ext' inserita correttamente nel df_excel_normalized")

    except Exception as e:
        print("Errore durante l'aggiornamento delle colonne Country e Tecnology nel df_excel_normalized:", str(e))
        #self.log_manager.log(f"Errore durante l'aggiornamento della FL in df_excel_normalizedil: {str(e)}", "error")
        return False, None 
    
    # Salvo il df_excel_normalized aggiornato
    success = Excel_file_mng.save_excel_file_advanced(df_excel_normalized, "df_excel_normalized_updated.xlsx")

    # ----- Modifica df_IW29 ------
    # Aggiungo la colonna 'OdM_data_inizio_cardine' nel df_IW29 ricavandolo dal df_AFKO
    # Crea dizionario di lookup per una singola colonna
    lookup_dict = df_AFKO.set_index('Ordine')['Data inizio cardine'].to_dict()
    df_IW29['OdM_data_inizio_cardine'] = df_IW29['Ordine'].map(lookup_dict)

    # Aggiungo la colonna 'Stato_sistema' nel df_IW29 ricavandolo dal df_IW39
    # Crea dizionario di lookup per una singola colonna
    lookup_dict = df_IW39.set_index('Ordine')['Stato sistema'].to_dict()
    df_IW29['OdM_stato_sistema'] = df_IW29['Ordine'].map(lookup_dict)    

    # Aggiungo la colonna 'PlantName' nel df_IW29 ricavandolo dal df_plants
    # Crea dizionario di lookup per una singola colonna
    lookup_dict = df_plants.set_index('Functional Loc.')['Description'].to_dict()
    df_IW29['Plant_name'] = df_IW29['Sede tecnica'].str[:8].map(lookup_dict)

    # Aggiungo la colonna 'PlantName' nel df_IW29 ricavandolo dal df_plants
    # Crea dizionario di lookup per una singola colonna
    lookup_dict = df_plants.set_index('Functional Loc.')['Strategy Semplified'].to_dict()
    df_IW29['Strategy'] = df_IW29['Sede tecnica'].str[:8].map(lookup_dict)

    # Aggiungo la colonna 'Tecnologia' nel df_IW29
    # Prelevo il terzo carattere della colonna 'Sede tecnica'
    df_IW29['Tecnologia'] = df_IW29['Sede tecnica'].str[2]  # Indice 2 = terzo carattere

    # ----- Modifica df_IW39 ------
    # Aggiungo la colonna 'PlantName' nel df_IW39 ricavandolo dal df_plants
    # Crea dizionario di lookup per una singola colonna
    lookup_dict = df_plants.set_index('Functional Loc.')['Strategy Semplified'].to_dict()
    df_IW39['Strategy'] = df_IW39['Sede tecnica'].str[:8].map(lookup_dict)

    # Aggiungo la colonna 'Tecnologia' nel df_IW39
    # Prelevo il terzo carattere della colonna 'Sede tecnica'
    df_IW39['Tecnologia'] = df_IW39['Sede tecnica'].str[2]  # Indice 2 = terzo carattere   

    # Aggiungo la colonna 'PlantName' nel df_IW39 ricavandolo dal df_plants
    # Crea dizionario di lookup per una singola colonna
    lookup_dict = df_plants.set_index('Functional Loc.')['Description'].to_dict()
    df_IW39['Plant_name'] = df_IW39['Sede tecnica'].str[:8].map(lookup_dict)

    # ----- Avvio l'elaborazione degli avvisi -----
    result, df_AdM = df_processor.process_iw29_with_ofa_actions(df_IW29, df_excel_normalized)
    if result == False:
        logger.info(f"Errore durante l'elaborazione del df IW29") # -> OK
        return
    
    # Salvo il df per utilizzarlo nella PowerBI
    success = Excel_file_mng.save_excel_file_advanced(df_AdM, "AdM_PowerBI.xlsx")

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
    result, df_OdM = df_processor.process_iw39_with_ofa_actions(df_IW39, df_excel_normalized)
    if result == False:
        logger.info(f"Errore durante l'elaborazione del df IW39") # -> OK
        return  

    # Salvo il df per utilizzarlo nella PowerBI
    success = Excel_file_mng.save_excel_file_advanced(df_OdM, "OdM_PowerBI.xlsx")

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
    
    return True

if __name__ == "__main__":
    # Esegui il test
    risultati = main()
    
    # I dataframe sono ora disponibili nel dizionario 'risultati'
    print(f"\nDataframe caricati: {list(risultati.keys())}")