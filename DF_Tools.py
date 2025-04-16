import pandas as pd
from collections import Counter
from typing import List, Dict, Optional
import os
import logging
import DF_Tools
from utils.decorators import error_logger

# Logger specifico per questo modulo
logger = logging.getLogger("DataFrameTools")

class DataFrameTools:
    """
    Classe di utility per la manipolazione dei DataFrame pandas
    """
    
    def __init__(self):
        logger.info("Inizializzazione DataFrameTools")
        self.df_utils = DF_Tools.DataFrameTools()


    # Controllo validità DataFrame
    @staticmethod
    def Add_Column_Check_ZPMR(df: pd.DataFrame):
        """
        Aggiunge una colonna al DataFrame per il controllo ZPMR.
        
        Args:
            df: DataFrame da modificare
        
        Returns:
            Boolean: True se l'operazione è riuscita, False altrimenti
        """
        # Controllo validità DataFrame
        if not isinstance(df, pd.DataFrame):
            raise TypeError("Il parametro df deve essere un DataFrame Pandas valido")
        
        # Controllo se il DataFrame è vuoto
        if df.empty:
            print("Il DataFrame è vuoto.")
            return False
        
        # Verifica l'esistenza della colonna 'FL' nel DataFrame
        required_column = 'FL'
        if required_column not in df.columns:
            print(f"Errore: La colonna '{required_column}' non esiste nel DataFrame.")
            return False
        
        try:
            # Crea una copia del DataFrame per evitare modifiche all'originale durante l'elaborazione
            df_temp = pd.DataFrame()
            df_temp['FL'] = df['FL'].copy()
            
            # Controllo se ci sono valori nulli nella colonna FL
            if df_temp['FL'].isna().any():
                print("Attenzione: La colonna 'FL' contiene valori nulli.")
            
            # Aggiungi il livello lunghezza e controlla errori
            df_temp, error_level = DataFrameTools.add_level_lunghezza(df_temp, 'FL')
            if error_level is not None:
                print(f"Errore nell'aggiunta del livello lunghezza: {error_level}")
                return False
            
            # Verifica l'esistenza delle colonne necessarie per la concatenazione
            required_columns = ["Livello_6", "Livello_5", "Livello_4", "Livello_3", "FL_Lunghezza"]
            missing_columns = [col for col in required_columns if col not in df_temp.columns]
            
            if missing_columns:
                print(f"Errore: Le seguenti colonne richieste non esistono: {', '.join(missing_columns)}")
                return False
            
            # Aggiunge la colonna Check per la verifica delle FL nelle tabelle globali
            df_temp, error_concat = DataFrameTools.add_concatenated_column_FL(
                df_temp, "Livello_6", "Livello_5", "Livello_4", "Livello_3", "FL_Lunghezza"
            )
            
            if error_concat is not None:
                print(f"Errore nella concatenazione delle colonne: {error_concat}")
                return False
            
            # Verifica che la colonna 'Check' sia stata creata correttamente
            if 'Check' not in df_temp.columns:
                print("Errore: La colonna 'Check' non è stata creata.")
                return False
                
            # Copia la colonna 'Check' nel DataFrame originale
            df['Check'] = df_temp['Check']
            print("Colonna 'Check' aggiunta con successo.")
            return True
            
        except Exception as e:
            print(f"Si è verificata un'eccezione durante l'elaborazione: {str(e)}")
            return False    

    @staticmethod
    @error_logger(logger=logger) 
    def load_dataframes_from_csv(df_names, base_path="", encoding="utf-8", separator=";"):
        """
        Verifica l'esistenza dei file CSV e carica i dati nei corrispettivi DataFrame.
        
        Args:
            df_names: Lista di nomi dei DataFrame da caricare
            base_path: Percorso base dove cercare i file CSV
            encoding: Encoding da utilizzare per leggere i file CSV
        
        Returns:
            dict: Dizionario contenente i DataFrame caricati
            
        Raises:
            FileNotFoundError: Se uno o più file non esistono
        """
        try:
            # Verifica che tutti i file esistano
            missing_files = []
            for df_name in df_names:
                # Ricavo i nomi dei file a partire dai DF
                file_name = f"{df_name}.csv"
                file_path = os.path.join(base_path, file_name)
                
                if not os.path.exists(file_path):
                    missing_files.append(file_name)
            
            # Se mancano file, solleva un'eccezione
            if missing_files:
                raise FileNotFoundError(f"I seguenti file non esistono: {', '.join(missing_files)}")
            
            # Tutti i file esistono, carica i dati
            dataframes = {}
            for df_name in df_names:
                # Ricavo i nomi dei file a partire dai DF
                file_name = f"{df_name}.csv"
                file_path = os.path.join(base_path, file_name)
                
                # Carica il file CSV nel DataFrame
                dataframes[df_name] = pd.read_csv(file_path, encoding=encoding, sep=separator)
                print(f"Caricato {file_name} in {df_name}")
            
            return dataframes
        
        except Exception as e:
            print(f"Errore durante il caricamento dei file CSV: {str(e)}")
            raise    

    @staticmethod
    @error_logger(logger=logger) 
    def save_dataframe_to_csv(df: pd.DataFrame, 
                            file_path: str,
                            separator: str = ";",
                            index: bool = False,
                            encoding: str = 'utf-8') -> bool:
        """
        Salva un DataFrame esistente in un file CSV con i dovuti controlli.
        
        Args:
            df: DataFrame Pandas da salvare
            file_path: Percorso completo del file CSV dove salvare i dati
            separator: Separatore da utilizzare nel file CSV (default: ";")
            index: Se includere l'indice del DataFrame nel file CSV (default: False)
            encoding: Codifica del file (default: 'utf-8')
        
        Returns:
            bool: True se l'operazione è completata con successo, False altrimenti
            
        Raises:
            TypeError: Se df non è un DataFrame valido
            ValueError: Se file_path è vuoto o non valido
            IOError: Se si verificano problemi nella scrittura del file
        """
        # Controllo validità del DataFrame
        if not isinstance(df, pd.DataFrame):
            raise TypeError("Il parametro df deve essere un DataFrame Pandas valido")
        
        # Controllo validità file_path
        if not isinstance(file_path, str) or not file_path.strip():
            raise ValueError("file_path deve essere una stringa valida")
        
        # Controllo validità separator
        if not isinstance(separator, str):
            raise TypeError("separator deve essere una stringa")
        
        # Controllo validità encoding
        if not isinstance(encoding, str):
            raise TypeError("encoding deve essere una stringa")
        
        try:
            # Crea la cartella di destinazione se non esiste
            import os
            os.makedirs(os.path.dirname(file_path), exist_ok=True)
            
            # Salva il DataFrame in un file CSV
            df.to_csv(file_path, 
                    sep=separator, 
                    index=index, 
                    encoding=encoding)
            
            # Verifica che il file sia stato creato
            if not os.path.exists(file_path):
                raise IOError(f"Il file {file_path} non è stato creato correttamente")
            
            # Verifica che il file abbia una dimensione maggiore di zero
            if os.path.getsize(file_path) == 0 and not df.empty:
                raise IOError(f"Il file {file_path} è stato creato ma è vuoto")
                
            return True
            
        except (IOError, PermissionError) as e:
            raise IOError(f"Errore durante il salvataggio del file CSV '{file_path}': {str(e)}")
        except Exception as e:
            # Rilancia qualsiasi altra eccezione con contesto aggiuntivo
            raise Exception(f"Errore imprevisto durante il salvataggio del DataFrame: {str(e)}") from e

    @staticmethod    
    @error_logger(logger=logger)
    def create_df_from_lists_ZPMR_CONTROL_FLn(header_string: str,
                                                lists_of_elements: list, 
                                                technology: str) -> pd.DataFrame:
            """
            Crea un DataFrame unico a partire da una lista di liste di elementi.
            Ogni lista rappresenta elementi per un diverso livello.
            
            Args:
                header_string: Stringa contenente i nomi delle colonne separati da delimitatore
                lists_of_elements: Lista di liste di elementi da inserire, ogni lista rappresenta un livello
                technology: Tecnologia utilizzata
            
            Returns:
                DataFrame unico combinato con tutte le liste di elementi
                
            Raises:
                ValueError: Se header_string è vuota o non contiene abbastanza colonne
                TypeError: Se lists_of_elements non è una lista di liste
                ValueError: Se tutte le liste in lists_of_elements sono vuote
                ValueError: Se technology o country_code non sono stringhe valide
            """
            # Controllo validità header_string
            if not isinstance(header_string, str):
                raise TypeError("header_string deve essere una stringa")
            if not header_string.strip():
                raise ValueError("header_string non può essere vuota")
            
            # Parsing dell'intestazione per ottenere i nomi delle colonne
            column_names = [col.strip() for col in header_string.split(';')]
            
            # Verifica che ci siano abbastanza colonne
            if (len(column_names) != 6):
                raise ValueError(f"L'intestazione deve contenere 6 colonne, trovate {len(column_names)}")
            
            # Controllo validità lists_of_elements
            if not isinstance(lists_of_elements, list):
                raise TypeError("lists_of_elements deve essere una lista di liste")
            
            # Verifico che almeno una lista non sia None e contenga elementi
            valid_lists = [sublist for sublist in lists_of_elements if sublist is not None]
            if not valid_lists or all(len(sublist) == 0 for sublist in valid_lists):
                raise ValueError("Tutte le liste in lists_of_elements sono vuote o None")
            
            # Controllo validità technology e country_code
            if not isinstance(technology, str) or not technology.strip():
                raise ValueError("technology deve essere una stringa valida")
            
            try:
                # Inizializzo una lista per raccogliere i DataFrame processati
                processed_dfs = []
                
                # Processo ogni lista di elementi con il relativo indice (livello)
                for index, elements_list in enumerate(lists_of_elements, start=3):
                    # Salto le liste None
                    if elements_list is None:
                        continue                    
                    # Verifico se la lista non è vuota
                    if len(elements_list) > 0:
                        # Creo un DataFrame vuoto con le colonne dell'intestazione
                        df = pd.DataFrame(columns=column_names)
                        
                        # Aggiungo gli elementi come righe
                        for element in elements_list:
                            # Creo una nuova riga con valori None
                            new_row = {col: None for col in df.columns}

                            # Imposto i valori nei campi specifici
                            if len(df.columns) == 6:
                                new_row[df.columns[0]] = "Z-R" + technology + "S"
                                new_row[df.columns[1]] = technology.strip()
                                new_row[df.columns[2]] = str(index).strip()  # Livello come stringa
                                new_row[df.columns[3]] = element.strip() if isinstance(element, str) else str(element)
                            
                            # Aggiungo la riga al DataFrame
                            df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
                        
                        # Aggiungo il DataFrame creato alla lista dei DataFrame processati
                        processed_dfs.append(df)
                
                # Combina tutti i DataFrame in uno solo
                if processed_dfs:
                    return pd.concat(processed_dfs, ignore_index=True)
                else:
                    return pd.DataFrame(columns=column_names)  # Restituisco un DataFrame vuoto con le colonne corrette
                    
            except Exception as e:
                # Rilanciamo l'eccezione con contesto aggiuntivo
                raise Exception(f"Errore nella creazione del DataFrame: {str(e)}") from e    
    
    @staticmethod    
    @error_logger(logger=logger)
    def create_df_from_elements_ZPMR_CTRL_ASS(header_string: str, elements_list: str, technology: str, country_code: str) -> pd.DataFrame:
        """
        Crea un DataFrame usando i nomi delle colonne dall'intestazione e gli elementi come righe
        
        Args:
            header_string: Stringa contenente i nomi delle colonne separati da delimitatore
            elements_list: Lista degli elementi da inserire, uno per riga
            technology: Tecnologia utilizzata
            country_code: Codice paese
        
        Returns:
            DataFrame con colonne definite dall'intestazione e gli elementi nelle righe
            
        Raises:
            ValueError: Se header_string è vuota o non contiene abbastanza colonne
            TypeError: Se elements_list non è una lista o stringa
            ValueError: Se elements_list è vuota
            ValueError: Se technology o country_code non sono stringhe valide
        """
        # Controllo validità header_string
        if not isinstance(header_string, str):
            raise TypeError("header_string deve essere una stringa")
        if not header_string.strip():
            raise ValueError("header_string non può essere vuota")
        
        # Parsing dell'intestazione per ottenere i nomi delle colonne
        column_names = [col.strip() for col in header_string.split(';')]
        
        # Verifica che ci siano abbastanza colonne
        if len(column_names) != 6:
            raise ValueError(f"L'intestazione deve contenere 6 colonne, trovate {len(column_names)}")
        
        # Controllo validità elements_list
        if not isinstance(elements_list, (list, str)):
            raise TypeError("elements_list deve essere una lista o una stringa")
        
        # Se elements_list è una stringa, convertiamola in lista
        if isinstance(elements_list, str):
            elements_list = [elem.strip() for elem in elements_list.split(',') if elem.strip()]
        
        if len(elements_list) == 0:
            raise ValueError("elements_list non può essere vuota")
        
        # Controllo validità technology e country_code
        if not isinstance(technology, str) or not technology.strip():
            raise ValueError("technology deve essere una stringa valida")
        if not isinstance(country_code, str) or not country_code.strip():
            raise ValueError("country_code deve essere una stringa valida")
        
        try:
            # Creo un DataFrame vuoto con le colonne dell'intestazione
            df = pd.DataFrame(columns=column_names)
            
            # Aggiungo gli elementi come righe
            for element in elements_list:
                # Creo una nuova riga con valori None
                new_row = {col: None for col in df.columns}
                
                # Imposto i valori nei campi specifici
                if len(df.columns) == 6:
                    new_row[df.columns[0]] = "Z-R" + technology + "S"
                    new_row[df.columns[1]] = technology.strip()
                    new_row[df.columns[2]] = "" #f_level.strip()
                    new_row[df.columns[3]] = element.strip()
                
                # Aggiungo la riga al DataFrame
                df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
            
            return df
        
        except Exception as e:
            # Rilanciamo l'eccezione con contesto aggiuntivo
            raise Exception(f"Errore nella creazione del DataFrame: {str(e)}") from e
    
    @staticmethod    
    @error_logger(logger=logger)
    def create_df_from_lists_ZPMR_CONTROL_FL2(header_string: str,
                                                lists_of_elements: list, 
                                                technology: str, 
                                                country_code: str) -> pd.DataFrame:
            """
            Crea un DataFrame unico a partire da una lista di liste di elementi.
            Ogni lista rappresenta elementi per un diverso livello.
            
            Args:
                header_string: Stringa contenente i nomi delle colonne separati da delimitatore
                lists_of_elements: Lista di liste di elementi da inserire, ogni lista rappresenta un livello
                technology: Tecnologia utilizzata
                country_code: Codice paese
            
            Returns:
                DataFrame unico combinato con tutte le liste di elementi
                
            Raises:
                ValueError: Se header_string è vuota o non contiene abbastanza colonne
                TypeError: Se lists_of_elements non è una lista di liste
                ValueError: Se tutte le liste in lists_of_elements sono vuote
                ValueError: Se technology o country_code non sono stringhe valide
            """
            # Controllo validità header_string
            if not isinstance(header_string, str):
                raise TypeError("header_string deve essere una stringa")
            if not header_string.strip():
                raise ValueError("header_string non può essere vuota")
            
            # Parsing dell'intestazione per ottenere i nomi delle colonne
            column_names = [col.strip() for col in header_string.split(';')]
            
            # Verifica che ci siano abbastanza colonne
            if (len(column_names) != 5):
                raise ValueError(f"L'intestazione deve contenere 5 colonne, trovate {len(column_names)}")
            
            # Controllo validità lists_of_elements
            if not isinstance(lists_of_elements, list):
                raise TypeError("lists_of_elements deve essere una lista di liste")
            
            # Verifico che almeno una lista non sia None e contenga elementi
            valid_lists = [sublist for sublist in lists_of_elements if sublist is not None]
            if not valid_lists or all(len(sublist) == 0 for sublist in valid_lists):
                raise ValueError("Tutte le liste in lists_of_elements sono vuote o None")
            
            # Controllo validità technology e country_code
            if not isinstance(technology, str) or not technology.strip():
                raise ValueError("technology deve essere una stringa valida")
            if not isinstance(country_code, str) or not country_code.strip():
                raise ValueError("country_code deve essere una stringa valida")
            
            try:
                # Inizializzo una lista per raccogliere i DataFrame processati
                processed_dfs = []
                
                # Processo ogni lista di elementi con il relativo indice (livello)
                for index, elements_list in enumerate(lists_of_elements, start=1):
                    # Salto le liste None
                    if elements_list is None:
                        continue

                    # Verifico se la lista non è vuota
                    if len(elements_list) > 0:
                        # Creo un DataFrame vuoto con le colonne dell'intestazione
                        df = pd.DataFrame(columns=column_names)
                        
                        # Aggiungo gli elementi come righe
                        for element in elements_list:
                            # Creo una nuova riga con valori None
                            new_row = {col: None for col in df.columns}
                            
                            # Imposto i valori nei campi specifici
                            if len(df.columns) == 5:
                                new_row[df.columns[0]] = "Z-R" + technology + "M"
                                new_row[df.columns[1]] = technology.strip()
                                new_row[df.columns[2]] = str(index).strip()  # Livello come stringa
                                new_row[df.columns[3]] = country_code.strip()
                                new_row[df.columns[4]] = element.strip() if isinstance(element, str) else str(element)
                            
                            # Aggiungo la riga al DataFrame
                            df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
                        
                        # Aggiungo il DataFrame creato alla lista dei DataFrame processati
                        processed_dfs.append(df)
                
                # Combina tutti i DataFrame in uno solo
                if processed_dfs:
                    return pd.concat(processed_dfs, ignore_index=True)
                else:
                    return pd.DataFrame(columns=column_names)  # Restituisco un DataFrame vuoto con le colonne corrette
                    
            except Exception as e:
                # Rilanciamo l'eccezione con contesto aggiuntivo
                raise Exception(f"Errore nella creazione del DataFrame: {str(e)}") from e

    @staticmethod    
    @error_logger(logger=logger)
    def stampa_risultato_differenze(risultato):
        """
        Stampa il numero di elementi e gli elementi presenti nel risultato
        della funzione trova_differenze.
        
        Parameters:
        -----------
        risultato : list o None
            Il risultato della funzione trova_differenze
        
        Returns:
        --------
        None
            Questa funzione non restituisce valori ma stampa a schermo il risultato
        """
        try:
            # Controlla se il risultato è None
            if risultato is None:
                print("Tutti gli elementi sono presenti nella tabella.")
                return
            
            # Controlla se il risultato è una lista
            if not isinstance(risultato, list):
                raise TypeError("Il risultato deve essere una lista o None")
            
            # Stampa il numero di elementi
            num_elementi = len(risultato)
            print(f"Numero di elementi mancanti nella tabella: {num_elementi}")
            
            # Stampa gli elementi
            if num_elementi > 0:
                print("Elementi:")
                for i, elemento in enumerate(risultato, 1):
                    print(f"{i}. {elemento}")
        
        except Exception as e:
            print(f"Errore durante la stampa del risultato: {str(e)}")    

    @staticmethod
    @error_logger(logger=logger)
    def trova_differenze(df1, df2, col1, col2):
        """
        Trova gli elementi nella prima colonna che non sono presenti nella seconda.
        
        Parameters:
        -----------
        df1 : pandas.DataFrame
            Il primo dataframe
        df2 : pandas.DataFrame
            Il secondo dataframe
        col1 : str
            Nome della colonna nel primo dataframe
        col2 : str
            Nome della colonna nel secondo dataframe
        
        Returns:
        --------
        list o None
            Lista di elementi presenti in df1[col1] ma non in df2[col2],
            o None se tutti gli elementi sono presenti
            
        Raises:
        -------
        TypeError
            Se gli input non sono del tipo corretto
        ValueError
            Se i dataframe sono vuoti o le colonne non esistono
        """
        try:
            # Verifiche sui tipi
            if not isinstance(df1, pd.DataFrame) or not isinstance(df2, pd.DataFrame):
                raise TypeError("Gli input devono essere pandas DataFrame")
            if not isinstance(col1, str) or not isinstance(col2, str):
                raise TypeError("I nomi delle colonne devono essere stringhe")
            
            # Verifica dataframe vuoti
            if df1.empty:
                raise ValueError("Il primo dataframe è vuoto")
            if df2.empty:
                raise ValueError("Il secondo dataframe è vuoto")
            
            # Verifica esistenza colonne
            if col1 not in df1.columns:
                raise ValueError(f"La colonna '{col1}' non esiste nel primo dataframe")
            if col2 not in df2.columns:
                raise ValueError(f"La colonna '{col2}' non esiste nel secondo dataframe")
            
            # Estrazione dei valori unici, escludendo stringhe vuote e valori null
            valori_col1 = {x for x in df1[col1].unique() if pd.notna(x) and (not isinstance(x, str) or x.strip() != '')}
            valori_col2 = {x for x in df2[col2].unique() if pd.notna(x) and (not isinstance(x, str) or x.strip() != '')}
            
            # Trova elementi in col1 ma non in col2
            differenze = valori_col1 - valori_col2
            
            # Restituisci risultato
            if not differenze:
                return None
            return sorted(list(differenze))  # Converti in lista ordinata per output più leggibile
        
        except Exception as e:
            # Cattura errori non previsti
            raise Exception(f"Errore durante il confronto delle colonne: {str(e)}")

    @staticmethod
    def pivot_hierarchy(df, values_col, level_col):
        """
        Trasforma un dataframe pivottando i livelli gerarchici in colonne.
        Applica strip() su tutti i valori stringa.
        
        Parameters:
        -----------
        df : pandas.DataFrame
            Il dataframe di input
        values_col : str
            Nome della colonna contenente i valori da pivottare
        level_col : str
            Nome della colonna contenente i livelli gerarchici
        
        Returns:
        --------
        pandas.DataFrame
            Dataframe trasformato con i livelli come colonne
            
        Raises:
        -------
        ValueError
            Se il dataframe è vuoto o se le colonne richieste non esistono
        TypeError
            Se gli input non sono del tipo corretto
        """
        try:
            # Verifica che gli input siano del tipo corretto
            if not isinstance(df, pd.DataFrame):
                raise TypeError("L'input 'df' deve essere un pandas DataFrame")
            if not isinstance(values_col, str) or not isinstance(level_col, str):
                raise TypeError("I nomi delle colonne devono essere stringhe")
                
            # Verifica che il dataframe non sia vuoto
            if df.empty:
                raise ValueError("Il dataframe è vuoto")
                
            # Verifica che le colonne esistano nel dataframe
            if values_col not in df.columns:
                raise ValueError(f"La colonna '{values_col}' non esiste nel dataframe")
            if level_col not in df.columns:
                raise ValueError(f"La colonna '{level_col}' non esiste nel dataframe")
                
            # Crea una copia del dataframe per non modificare l'originale
            df_clean = df.copy()
            
            # Applica strip() ai valori stringa
            if df_clean[values_col].dtype == 'object':
                df_clean[values_col] = df_clean[values_col].apply(lambda x: x.strip() if isinstance(x, str) else x)
                
            # Verifica che ci siano dati validi
            if df_clean[level_col].isna().all() or df_clean[values_col].isna().all():
                raise ValueError("Le colonne contengono solo valori nulli")
                
            # Procedi con la trasformazione
            max_rows = max(df_clean[level_col].value_counts())
            result = pd.DataFrame(index=range(max_rows))
            
            for level in sorted(df_clean[level_col].unique()):
                values = df_clean[df_clean[level_col] == level][values_col].values
                col_name = f'Livello_{level}'
                result[col_name.strip()] = pd.Series(values)
                
            return result
            
        except Exception as e:
            # Cattura eventuali altri errori non previsti
            raise Exception(f"Errore durante l'elaborazione del dataframe: {str(e)}")
    

    @staticmethod
    def get_last_char(df, colonna):
        """
        Restituisce i primi due caratteri di una colonna contenente dati univoci
        
        Args:
            df: DataFrame da verificare
            colonna: nome della colonna in cui analizzare il dato
            
        Returns:
            none : se sono presenti solo valori nulli
            i primi due caratteri della colonna indicata
        """        
        try:
            # Prende il primo valore non nullo
            primo_valore = df[colonna].dropna().iloc[0]
            if pd.notna(primo_valore):  # verifica che non sia nullo
                return primo_valore[2:3]
            else:
                return None
        except Exception as e:
            print(f"Errore nell'elaborazione della colonna {colonna}: {str(e)}")
            return None
        
    @staticmethod
    def get_first_two_chars(df, colonna):
        """
        Restituisce i primi due caratteri di una colonna contenente dati univoci
        
        Args:
            df: DataFrame da verificare
            colonna: nome della colonna in cui analizzare il dato
            
        Returns:
            none : se sono presenti solo valori nulli
            i primi due caratteri della colonna indicata
        """        
        try:
            # Prende il primo valore non nullo
            primo_valore = df[colonna].dropna().iloc[0]
            if pd.notna(primo_valore):  # verifica che non sia nullo
                return primo_valore[:2]
            else:
                return None
        except Exception as e:
            print(f"Errore nell'elaborazione della colonna {colonna}: {str(e)}")
            return None
        
    @staticmethod
    def check_dataframe(df, name="DataFrame"):
        """
        Esegue un controllo completo su un DataFrame
        
        Args:
            df: DataFrame da verificare
            name: Nome del DataFrame per i messaggi di errore
            
        Returns:
            bool: True se il DataFrame è valido
        """
        try:
            # Verifica se è None
            if df is None:
                print(f"{name} è None")
                return False
                
            # Verifica se è un DataFrame
            if not isinstance(df, pd.DataFrame):
                print(f"{name} non è un DataFrame valido")
                return False
                
            # Verifica se è vuoto
            if df.empty:
                print(f"{name} è vuoto")
                return False
                
            # Verifica se ha righe e colonne
            if df.shape[0] == 0 or df.shape[1] == 0:
                print(f"{name} non ha righe o colonne")
                return False
                
            return True
            
        except Exception as e:
            print(f"Errore durante la verifica di {name}: {str(e)}")
            return False

    @staticmethod
    def clean_data(data: pd.DataFrame) -> pd.DataFrame:
        """
        Legge i dati dalla clipboard, rimuove le righe di separazione e le colonne vuote,
        e gestisce le intestazioni duplicate.
        
        Returns:
            DataFrame Pandas pulito o None in caso di errore
        """
        try:
            if not data:
                print("Nessun dato trovato nella clipboard")
                return None

            # Divide in righe
            lines = data.strip().split('\n')
            
            # Filtra le righe, escludendo quelle che contengono solo trattini
            clean_lines = []
            for line in lines:
                # Rimuove spazi bianchi iniziali e finali
                line = line.strip()
                # Verifica se la riga è composta solo da trattini
                if line and not all(c == '-' for c in line.replace(' ', '')):
                    clean_lines.append(line)

            if not clean_lines:
                print("Nessuna riga valida trovata dopo la pulizia")
                return None

            # Dividi le righe in colonne usando il tab come separatore
            data_rows = [line.split('|') for line in clean_lines]
            
            # Prendi la prima riga come header
            original_headers = [header.strip() for header in data_rows[0]]
            
            # Gestisci gli header duplicati
            unique_headers = DataFrameTools.handle_duplicate_headers(original_headers)
            
            # Se sono stati trovati duplicati, stampalo
            duplicates = [header for header, count in Counter(original_headers).items() if count > 1]
            if duplicates:
                print("\nTrovate colonne con nomi duplicati:")
                for dup in duplicates:
                    print(f"- '{dup}' (rinominate con postfissi numerici)")

            # Crea il DataFrame con i nuovi header
            df = pd.DataFrame(data_rows[1:], columns=unique_headers)

            # Rimuove le colonne completamente vuote
            df = df.dropna(axis=1, how='all')
            
            # Rimuove le colonne dove tutti i valori sono stringhe vuote
            df = df.loc[:, ~(df == '').all()]
            
            # Reset dell'indice
            df = df.reset_index(drop=True)
         
            return df

        except Exception as e:
            print(f"Errore durante la pulizia dei dati: {str(e)}")
            return None
    
    @staticmethod        
    def handle_duplicate_headers(headers: List[str]) -> List[str]:
        """
        Gestisce le intestazioni duplicate aggiungendo un postfisso numerico
        
        Args:
            headers: Lista delle intestazioni originali
            
        Returns:
            Lista delle intestazioni con postfissi per i duplicati
        """
        # Conta le occorrenze di ogni header
        header_counts = Counter()
        unique_headers = []
        
        for header in headers:
            # Se l'header è già stato visto
            if header in header_counts:
                # Incrementa il contatore e aggiungi il postfisso
                header_counts[header] += 1
                unique_headers.append(f"{header}_{header_counts[header]}")
            else:
                # Prima occorrenza dell'header
                header_counts[header] = 0
                unique_headers.append(header)
        
        return unique_headers        


    @staticmethod
    @error_logger(logger=logger) 
    def add_level_lunghezza(df: pd.DataFrame, col1: str) -> pd.DataFrame:
        """
        Divide i valori di una colonna usando '-' come separatore e crea 6 nuove colonne con i risultati.
        
        Args:
            df: DataFrame da elaborare
            col1: Nome della colonna da splittare
        
        Returns:
            DataFrame con le nuove colonne aggiunte
            
        Raises:
            TypeError: Se df non è un DataFrame valido
            ValueError: Se col1 non è presente nel DataFrame
            ValueError: Se col1 contiene valori non-stringa che non possono essere splittati
        """
        # Controllo che df sia un DataFrame
        if not isinstance(df, pd.DataFrame):
            raise TypeError("Il parametro df deve essere un DataFrame pandas")
        
        # Controllo che df non sia vuoto
        if df.empty:
            return df.copy()  # Restituisce una copia del DataFrame vuoto
        
        # Controllo che col1 sia una stringa
        if not isinstance(col1, str):
            raise TypeError(f"Il nome della colonna deve essere una stringa, ricevuto {type(col1)}")
        
        # Controllo che col1 esista nel DataFrame
        if col1 not in df.columns:
            raise ValueError(f"La colonna '{col1}' non è presente nel DataFrame. "
                            f"Colonne disponibili: {', '.join(df.columns)}")
        
        # Crea una copia per non modificare l'originale
        result_df = df.copy()
        
        # Controllo valori non-stringa nella colonna
        non_string_mask = ~result_df[col1].apply(lambda x: isinstance(x, str))
        if non_string_mask.any():
            # Converti i valori non-stringa in stringhe
            result_df[col1] = result_df[col1].astype(str)
        
        # Funzione per splittare e padare a 6 elementi
        def split_and_pad(x):
            try:
                parts = x.split('-')
                # Converte ogni elemento in stringa e rimuove spazi
                parts = [str(part).strip() for part in parts]
                # Estende la lista a 6 elementi aggiungendo stringhe vuote
                parts.extend([''] * (6 - len(parts)))
                return pd.Series(parts[:6])
            except Exception as e:
                # Gestione di qualsiasi errore nello splitting
                raise ValueError(f"Errore nello splitting del valore '{x}': {str(e)}")
        
        try:
            # Crea le colonne numerate da 1 a 6
            result_df[['Livello_1', 'Livello_2', 'Livello_3', 'Livello_4', 'Livello_5', 'Livello_6']] = result_df[col1].apply(split_and_pad)
            # Aggiunge la colonna FL_Lunghezza con il numero di elementi dopo lo split
            result_df['FL_Lunghezza'] = result_df['FL'].apply(lambda x: len(x.split('-')))

        except Exception as e:
            raise ValueError(f"Errore nella creazione delle colonne di livello: {str(e)}")
        
        return result_df

    @staticmethod
    @error_logger(logger=logger) 
    def add_concatenated_column_FL(df: pd.DataFrame, 
                            col1: str, 
                            col2: str, 
                            col3: str, 
                            col4: str,
                            col5: str,
                            new_column_name: str = 'Check',
                            separator: str = '_') -> pd.DataFrame:
        """
        Aggiunge una nuova colonna che concatena i valori di 4 colonne specificate usando un separatore
        
        Args:
            df: DataFrame di input
            col1: Nome della prima colonna -> Valore Livello
            col2: Nome della seconda colonna -> Valore Liv. Superiore
            col3: Nome della terza colonna -> Valore Liv. Superiore_1
            col4: Nome della quarta colonna -> colonna contenente la lunghezza della FL
            new_column_name: Nome della nuova colonna da creare (default: 'Check')
            separator: Carattere/i da usare come separatore (default: '_')
                
        Returns:
            DataFrame con la nuova colonna contenente i valori concatenati
        """
        # Verifica che tutte le colonne esistano nel DataFrame
        required_cols = [col1, col2, col3, col4, col5]
        if not all(col in df.columns for col in required_cols):
            raise ValueError("Una o più colonne specificate non esistono nel DataFrame")

        def create_concatenated_value(row):
            # Verifica che col4 non sia nullo
            if (str(row[col3]).strip(' \t\n\r') == ""):
                return None
            # add_concatenated_column(df, "Livello_6", "Livello_5", "Livello_4",  "Livello_3", "FL_Lunghezza")
            #                               col1        col2         col3           col4        col5
            if (row[col1].strip(' \t\n\r') != ""): # se è presente il 6 livello allora concateno 6-5-4-Lunghezza
                result = f"{str(row[col1].strip(' \t\n\r'))}{separator}{str(row[col2].strip(' \t\n\r'))}{separator}{str(row[col3].strip(' \t\n\r'))}{separator}{str(str(row[col5]).strip(' \t\n\r'))}"
                return result
            elif (row[col2].strip(' \t\n\r') != ""): # se è presente il 5 livello allora concateno 5-4-3-Lunghezza
                result = f"{str(row[col2].strip(' \t\n\r'))}{separator}{str(row[col3].strip(' \t\n\r'))}{separator}{str(str(row[col4]).strip(' \t\n\r'))}{separator}{str(str(row[col5]).strip(' \t\n\r'))}"
                return result
            elif (row[col3].strip(' \t\n\r') != ""):
                result = f"{str(row[col3].strip(' \t\n\r'))}{separator}{str(str(row[col4]).strip(' \t\n\r'))}{separator}{str(str(row[col5]).strip(' \t\n\r'))}"
                return result
            else:
                return None                 
        
        df_copy = df.copy()
        df_copy[new_column_name] = df_copy.apply(create_concatenated_value, axis=1)
        return df_copy
    
    @staticmethod
    def add_concatenated_column_SAP(df: pd.DataFrame, 
                            col1: str, 
                            col2: str, 
                            col3: str, 
                            col4: str,
                            new_column_name: str = 'Check',
                            separator: str = '_') -> pd.DataFrame:
        """
        Aggiunge una nuova colonna che concatena i valori di 4 colonne specificate usando un separatore
        
        Args:
            df: DataFrame di input
            col1: Nome della prima colonna -> Valore Livello
            col2: Nome della seconda colonna -> Valore Liv. Superiore
            col3: Nome della terza colonna -> Valore Liv. Superiore_1
            col4: Nome della quarta colonna -> colonna contenente la lunghezza della FL
            new_column_name: Nome della nuova colonna da creare (default: 'Check')
            separator: Carattere/i da usare come separatore (default: '_')
                
        Returns:
            DataFrame con la nuova colonna contenente i valori concatenati
        """
        # Verifica che tutte le colonne esistano nel DataFrame
        required_cols = [col1, col2, col3, col4]
        if not all(col in df.columns for col in required_cols):
            raise ValueError("Una o più colonne specificate non esistono nel DataFrame")

        def create_concatenated_value(row):
            if (row[col3].strip(' \t\n\r') != ""):
                result = f"{str(row[col1].strip(' \t\n\r'))}{separator}{str(row[col2].strip(' \t\n\r'))}{separator}{str(row[col3].strip(' \t\n\r'))}{separator}{str(str(row[col4]).strip(' \t\n\r'))}"
                return result
            elif (row[col2].strip(' \t\n\r') != ""):
                result = f"{str(row[col1].strip(' \t\n\r'))}{separator}{str(row[col2].strip(' \t\n\r'))}{separator}{str(str(row[col4]).strip(' \t\n\r'))}"
                return result
            elif (row[col1].strip(' \t\n\r') != ""):
                result = f"{str(row[col1].strip(' \t\n\r'))}{separator}{str(str(row[col4]).strip(' \t\n\r'))}"
                return result
            else:
                return None    
        
        df_copy = df.copy()
        df_copy[new_column_name] = df_copy.apply(create_concatenated_value, axis=1)
        return df_copy    
    
    @staticmethod
    def analyze_data(df: pd.DataFrame, df_name: str = '') -> None:
        """
        Analizza il DataFrame pulito e mostra informazioni utili
        
        Args:
            df: DataFrame da analizzare
        """
        print("\nAnalisi del DataFrame pulito:" + df_name)
        print(f"Dimensioni: {df.shape}")
        print("\nColonne presenti:")
        for col in df.columns:
            print(f"- '{col}'")
        print("\nPrime 5 righe:")
        print(df.head())
    
    @staticmethod
    def strip_column_headers(df: pd.DataFrame) -> pd.DataFrame:
        """
        Rimuove gli spazi iniziali e finali dai nomi delle colonne
        
        Args:
            df: DataFrame di input
            
        Returns:
            DataFrame con i nomi delle colonne puliti
        """
        df_copy = df.copy()
        df_copy.columns = df_copy.columns.str.strip()
        return df_copy
    
    @staticmethod
    def combine_columns(df: pd.DataFrame, col1: str, col2: str, new_col: str, separator: str = '-') -> pd.DataFrame:
        """
        Crea una nuova colonna combinando due colonne esistenti
        
        Args:
            df: DataFrame di input
            col1: Nome della prima colonna
            col2: Nome della seconda colonna
            new_col: Nome della nuova colonna
            separator: Separatore da usare nella concatenazione
            
        Returns:
            DataFrame con la nuova colonna combinata
        """
        df_copy = df.copy()
        df_copy[new_col] = df_copy[col1] + separator + df_copy[col2]
        return df_copy