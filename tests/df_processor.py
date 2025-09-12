import pandas as pd
import os
from pathlib import Path
import logging

# Configurazione logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

class DfProcessor:

    def process_iw29_with_ofa_actions(df_IW29, df_excel_normalized) -> tuple[bool, pd.DataFrame | None]:
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
            return False
        
        if df_excel_normalized is None or df_excel_normalized.empty:
            logger.error("df_excel_normalized è vuoto o None")
            return False
        
        # Verifica colonne necessarie
        if 'Avviso' not in df_IW29.columns:
            logger.error("Colonna 'Avviso' non trovata in df_IW29")
            return False
        
        if 'Ordine' not in df_IW29.columns:
            logger.error("Colonna 'Ordine' non trovata in df_IW29")
            return False
        
        if 'AdM' not in df_excel_normalized.columns or 'action' not in df_excel_normalized.columns:
            logger.error("Colonne 'AdM' o 'action' non trovate in df_excel_normalized")
            return False
            
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

        # Creo un dizionario con le maschere per le diverse azioni
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
        
        # Processa ogni colonna del df_result realizzando e applicando le maschere e i flitri sulle date secondo le azioni in OFA
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
                if 'Data cr.' in df_result.columns:
                    # Definisci il range di date
                    data_inizio = pd.to_datetime('01.06.2025', format='%d.%m.%Y')  # Le tue date
                    data_fine = pd.to_datetime('30.06.2025', format='%d.%m.%Y')
                    
                    # Maschera per il range di date (estremi compresi) - NO conversione necessaria
                    mask_date_range = (df_result['Data cr.'] >= data_inizio) & \
                                    (df_result['Data cr.'] <= data_fine)
                else:
                    print(f"ERRORE nella valutazione della colonna 'Data'")
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
                    print(f"ERRORE nella valutazione della colonna 'Ordine'")
                    continue                

                # Condizione Avviso chiuso
                if 'St.sist.' in df_result.columns:
                    mask_meco = df_result['St.sist.'].str.contains('MECO', na=False, case=False)
                    count = mask_meco.sum()   
                else:
                    print(f"ERRORE nella valutazione della colonna 'St.sist.'")
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
        # Il denominatore assume valore 1 se il numeratore è 1 oppure se
        # la colonna 'Ordine' != "" E 
        # la colonna data modifica 'Mod. il' è vuota E
        # la data di inizio cardine è contenuta nel range di date selezionato nella GUI
        # altrimenti = 0
        
        # Costruisco la maschera a partire dai valori del numeratore
        mask_numeratore = df_result['Numeratore']

        # Definisco condizione per verificare che l'AdM abbia il campo 'OdM_data_inizio_cardine' entro il range di date selezionato nella GUI
        if 'OdM_data_inizio_cardine' in df_result.columns:
            # Definisci il range di date
            data_inizio = pd.to_datetime('01.06.2025', format='%d.%m.%Y')  # Le tue date
            data_fine = pd.to_datetime('30.06.2025', format='%d.%m.%Y')
            
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

        
        return True, df_result

    def process_iw39_with_ofa_actions(df_IW39, df_excel_normalized) -> tuple[bool, pd.DataFrame | None]:
        """
        Processa il DataFrame df_IW39 aggiungendo colonne OFA basate sui dati di df_excel_normalized
        
        Args:
            df_IW39 (pd.DataFrame): DataFrame IW39
            df_excel_normalized (pd.DataFrame): DataFrame normalizzato con colonne 'AdM', 'OdM' e 'action'
        
        Returns:
            pd.DataFrame: DataFrame df_IW39 aggiornato con le nuove colonne OFA
            
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
            return False
        
        if df_excel_normalized is None or df_excel_normalized.empty:
            logger.error("df_excel_normalized è vuoto o None")
            return False
        
        if 'Ordine' not in df_IW39.columns:
            logger.error("Colonna 'Ordine' non trovata in df_IW39")
            return False
        
        if 'OdM' not in df_excel_normalized.columns or 'action' not in df_excel_normalized.columns:
            logger.error("Colonne 'OdM' o 'action' non trovate in df_excel_normalized")
            return False
            
        # Crea una copia del DataFrame per non modificare l'originale
        df_result = df_IW39.copy()
        
        # Definisci le nuove colonne OFA
        ofa_columns = ['Creati_in_OFA', 'Visualizzati_in_OFA', 'Modificati_in_OFA', 'Operazione_gestita_in_OFA',
                        'Aggiunto_materiale', 'Gestione_materiale', 'Allegati_in_OFA', 'Conferma_ore_in_OFA',
                        'Chiusi_in_OFA']
        
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
            df_action = df_excel_normalized[df_excel_normalized['action'] == action] # Creo un df filtrando df_excel_normalized in base all'azione considerata
            print(f"Dimensioni del df per action: {action} = {df_action.shape[0]}")
            if df_action.empty:
                logger.warning(f"Nessun OdM per azione '{action}'")
                mask_dict[action] = pd.Series([False] * len(df_result), index=df_result.index)  # Mask vuota
                continue
            # Verifica che ci siano OdM validi per ogni azione
            valid_odm = df_action['OdM'].dropna()
            if valid_odm.empty:
                logger.warning(f"Nessun OdM valido per azione '{action}'")
                mask_dict[action] = pd.Series([False] * len(df_result), index=df_result.index)
                continue
                # Mask base
            mask = df_result['Ordine'].isin(valid_odm)
            mask_dict[action] = mask
        
        # Processa ogni colonna del df_result realizzando e applicando le maschere  secondo le azioni in OFA
        for target_column in ofa_columns:
            # Imposto un valore di default da utilizzare in caso di errore
            mask_finale = pd.Series([False] * len(df_result), index=df_result.index)

            # Condizioni speciali
            if target_column == 'Creati_in_OFA': # -> OK
                mask_base = mask_dict['CREATE WORKORDER']
                # Aggiunge condizione Sis.Legacy
                if 'Sis Legacy' in df_result.columns:
                    mask_legacy = df_result['Sis Legacy'].str.contains('OFA', na=False, case=False) # da aggiungere filtro sulla data creazione deve essere nel mese corrente
                else:
                    print(f"ERRORE nella valutazione della colonna 'Sis Legacy'")
                    continue

                # Aggiungi condizione per verificare che l'OdM abbia il campo data di acquisizione 'Data acq.' entro il range di date selezionato nella GUI
                if 'Data acq.' in df_result.columns:
                    # Definisci il range di date
                    data_inizio = pd.to_datetime('01.06.2025', format='%d.%m.%Y')  # Le tue date
                    data_fine = pd.to_datetime('30.06.2025', format='%d.%m.%Y')
                    
                    # Maschera per il range di date (estremi compresi) - NO conversione necessaria
                    mask_date_range = (df_result['Data acq.'] >= data_inizio) & \
                                    (df_result['Data acq.'] <= data_fine)
                else:
                    print(f"ERRORE nella valutazione della colonna 'Data acq.'")
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
                                    df_result['Operazione_gestita_in_OFA'] |                                
                                    df_result['Aggiunto_materiale'] |
                                    df_result['Gestione_materiale'] |
                                    df_result['Allegati_in_OFA'] |
                                    df_result['Conferma_ore_in_OFA'] |                                
                                    df_result['Chiusi_in_OFA'])  
        
        logger.info(f"Colonna {'Numeratore'}: {df_result['Numeratore'].sum()} True") # -> OK
        
        # Costruisco la colonna 'Denominatore'
        # Il denominatore assume valore 1 se il numeratore è 1 oppure se
        # la colonna 'Ordine' != "" E 
        # la colonna data modifica 'Mod. il' è vuota E
        # la data di inizio cardine è contenuta nel range di date selezionato nella GUI
        # altrimenti = 0
        
        # Costruisco la maschera a partire dai valori del numeratore
        mask_numeratore = df_result['Numeratore']

        # Definisco condizione per verificare che l'OdM abbia il campo data inizio cardine 'In. card.' entro il range di date selezionato nella GUI
        if 'In. card.' in df_result.columns:
            # Definisci il range di date
            data_inizio = pd.to_datetime('01.06.2025', format='%d.%m.%Y')  # Le tue date
            data_fine = pd.to_datetime('30.06.2025', format='%d.%m.%Y')
            
            # Maschera per il range di date (estremi compresi)
            mask_date_inizio_cardine = (df_result['In. card.'] >= data_inizio) & (df_result['In. card.'] <= data_fine)
        else:
            print(f"ERRORE nella valutazione della colonna 'In. card.'")
            return False
        
        # Definisco condizione per verificare che l'OdM abbia il campo data modifica 'Data mod.' entro il range di date selezionato nella GUI
        if 'Data mod.' in df_result.columns:
            # Definisci il range di date
            data_inizio = pd.to_datetime('01.06.2025', format='%d.%m.%Y')  # Le tue date
            data_fine = pd.to_datetime('30.06.2025', format='%d.%m.%Y')
            
            # Maschera per il range di date (estremi compresi)
            mask_date_modifica = (df_result['Data mod.'] >= data_inizio) & (df_result['Data mod.'] <= data_fine)
        else:
            print(f"ERRORE nella valutazione della colonna 'Data mod.'")
            return False    
        
        # Condizione tipo OdM diverso da "M1"
        if 'Tp.' in df_result.columns:
            mask_tipo_ordine = (df_result['Tp.'] != 'M1') & (df_result['Tp.'].notna()) # Escludo i tipo 'M1' e i valori NaN
        else:
            print(f"ERRORE nella valutazione della colonna 'Tp.'")
            return False

        # Condizione stato sistema OdM
        if 'Stato sistema' in df_result.columns:
            mask_stato_sistema = df_result['Stato sistema'].str.contains('TECO|CONC|RIL', na=False, case=False)
            
        else:
            print(f"ERRORE nella valutazione della colonna 'Stato sistema'")
            return False  

        df_result['Denominatore'] = False
        mask_finale = mask_numeratore | (mask_stato_sistema & mask_date_inizio_cardine & mask_date_modifica & mask_tipo_ordine)
        target_column = 'Denominatore'
        # Applica
        df_result.loc[mask_finale, target_column] = True
        logger.info(f"Colonna {target_column}: {df_result[target_column].sum()} True") # -> OK

        
        return True, df_result


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
    df_plants = dataframes.get('df_plants')

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
    result, df_AdM = process_iw29_with_ofa_actions(df_IW29, df_excel_normalized)
    if result == False:
        logger.info(f"Errore durante l'elaborazione del df IW29") # -> OK
        return
    
    # Salvo il df per utilizzarlo nella PowerBI
    success = loader.save_excel_file_advanced(df_AdM, "AdM_PowerBI.xlsx")

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
    result, df_OdM = process_iw39_with_ofa_actions(df_IW39, df_excel_normalized)
    if result == False:
        logger.info(f"Errore durante l'elaborazione del df IW39") # -> OK
        return  

    # Salvo il df per utilizzarlo nella PowerBI
    success = loader.save_excel_file_advanced(df_OdM, "OdM_PowerBI.xlsx")

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
    


    return dataframes


if __name__ == "__main__":
    # Esegui il test
    risultati = main()
    
    # I dataframe sono ora disponibili nel dizionario 'risultati'
    print(f"\nDataframe caricati: {list(risultati.keys())}")