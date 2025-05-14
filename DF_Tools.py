import pandas as pd
from collections import Counter
from typing import List, Dict, Optional
import os
import logging
from PyQt5.QtCore import QObject, pyqtSignal

# Logger specifico per questo modulo
logger = logging.getLogger("DataFrameTools")

class DataFrameTools:
    """
    Classe di utility per la manipolazione dei DataFrame pandas
    """
    
    def __init__(self):
        pass
        #logger.info("Inizializzazione DataFrameTools")  

    # Definizione aggiornata del segnale nella classe SAPDataExtractor
    logMessage = pyqtSignal(str, str, bool, bool, object, str, tuple, dict)
    # (message, level, update_status, update_log, min_display_seconds, origin, args, kwargs)

    def log(self, message, level='info', update_status=True, update_log=True, min_display_seconds=None, origin="DataFrameTools", *args, **kwargs):
        """
        Emette un segnale di log con controllo dell'origine
        
        Args:
            message: Il messaggio da registrare
            level: Livello del log ('info', 'warning', 'error', ecc.)
            update_status: Se aggiornare la statusBar
            update_log: Se aggiornare il widget di log visuale
            min_display_seconds: Tempo minimo di visualizzazione
            origin: Origine del messaggio per evitare duplicazioni ('main' o altro)
            *args, **kwargs: Argomenti per la formattazione del messaggio
        """
        # Registra nel logger locale
        logger_method = getattr(logger, level) if hasattr(logger, level) else logger.info
        
        # Formatta il messaggio per il logger
        try:
            if kwargs and not args:
                formatted_message = message.format(**kwargs)
            elif args:
                formatted_message = message % args
            else:
                formatted_message = message
        except Exception as e:
            formatted_message = f"{message} (Errore formattazione: {str(e)})"
        
        # Registra nel logger locale
        logger_method(formatted_message)
        
        # Emetti il segnale con l'informazione sull'origine
        # Passa origin come parametro aggiuntivo al segnale
        self.logMessage.emit(message, level, update_status, update_log, min_display_seconds, origin, args, kwargs)

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

    def clean_data(data: pd.DataFrame) -> pd.DataFrame:
        """
        Verifica e pulisce i dati prima di caricarli in un DataFrame.
        
        Returns:
            DataFrame Pandas pulito o None in caso di errore
        """
        try:
            if not data:
                print("DF non valido")
                return None

            # Divide in righe
            lines = data.strip().split('\n')
            
            # Filtra le righe, escludendo quelle che contengono solo trattini
            filtered_lines = []
            for line in lines:
                # Rimuove spazi bianchi iniziali e finali
                line = line.strip()
                # Verifica se la riga è composta solo da trattini
                if line and not all(c == '-' for c in line.replace(' ', '')):
                    filtered_lines.append(line)

            if not filtered_lines:
                print("Nessuna riga valida trovata dopo la pulizia")
                return None
            
            # La prima riga è l'intestazione
            original_headers = filtered_lines[0]
            list_original_headers = original_headers.strip().split('|')
            print(f"Intestazioni originali: {list_original_headers} - lunghezza: {len(list_original_headers)}")

            # Gestisci gli header duplicati
            list_unique_headers = DataFrameTools.handle_duplicate_headers(list_original_headers)
            print(f"Intestazioni uniche: {list_unique_headers} - lunghezza: {len(list_unique_headers)}")

            expected_pipes = original_headers.count('|')
            
            # Creo un array per memorizzare le righe corrette
            corrected_lines = []
            # Correggi le righe con occorrenze extra di pipe
            for line in filtered_lines[1:]:
                actual_pipes = line.count('|')
                
                if actual_pipes > expected_pipes:
                    # Calcola quanti pipe extra ci sono
                    extra_pipes = actual_pipes - expected_pipes
                    
                    # Trova l'indice del quarto pipe (che precede la colonna Descrizione)
                    pipe_indices = [i for i, char in enumerate(line) if char == '|']
                    if len(pipe_indices) >= 4:
                        description_start = pipe_indices[3] + 1
                        description_end = pipe_indices[4 + extra_pipes]
                        
                        # Estrai la descrizione
                        description = line[description_start:description_end]
                        
                        # Sostituisci i pipe nella descrizione con trattini
                        corrected_description = description.replace('|', '-')
                        print(f"Descrizione corretta: {corrected_description}")
                        
                        # Ricostruisci la riga
                        corrected_line = line[:description_start] + corrected_description + line[description_end:]
                        
                        # Verifica che la correzione abbia funzionato
                        if corrected_line.count('|') == expected_pipes:
                            corrected_lines.append(corrected_line)
                        else:
                            msg = (f"Impossibile correggere la riga: {line}")
                            # Se la correzione non ha funzionato, segnala l'errore
                            print(msg)
                            # Genera un exception
                            raise ValueError(msg)
                    else:
                        msg = (f"La riga non ha abbastanza separatori pipe: {line}")
                        # Se la correzione non ha funzionato, segnala l'errore
                        print(msg)
                        # Genera un exception
                        raise ValueError(msg)                                            
                else:
                    # La riga è già corretta o ha meno pipe del previsto
                    corrected_lines.append(line)
            
            # Ora che abbiamo corretto le righe, le convertiamo in DataFrame
            data_rows = []
            for line in corrected_lines:
                fields = [field.strip() for field in line.split('|')]
                # Rimuovi i campi vuoti all'inizio e alla fine (dovuti al pipe iniziale/finale)
                # fields = [field for field in fields if field]
                data_rows.append(fields)
            
            # Estrai gli header e i dati
            headers = list_unique_headers
            data = data_rows
            
            # Crea il DataFrame
            df = pd.DataFrame(data, columns=headers)     
                
            # Rimuoviamo le colonne senza intestazione
            unnamed_cols = [col for col in df.columns if col == '' or pd.isna(col) or str(col).startswith('Unnamed_')]
            df_cleaned = df.drop(columns=unnamed_cols)
            
            # Reset dell'indice
            df_cleaned = df_cleaned.reset_index(drop=True)

            
            return df_cleaned

        except Exception as e:
            print(f"Errore durante la pulizia dei dati: {str(e)}")
            return None
         
    def handle_duplicate_headers(headers: List[str]) -> List[str]:
        """
        Gestisce le intestazioni duplicate aggiungendo un postfisso numerico
        e rinomina le intestazioni vuote con il prefisso "Unnamed_"
        
        Args:
            headers: Lista delle intestazioni originali
            
        Returns:
            Lista delle intestazioni con postfissi per i duplicati e rinominate per le vuote
        """
        # Inizializziamo il contatore per le colonne senza nome
        unnamed_counter = 0
        
        # Per prima cosa, rinominiamo le intestazioni vuote
        processed_headers = []
        for header in headers:
            # Verifica se l'header è vuoto (dopo strip)
            if not header or header.strip() == "":
                # Assegna un nome "Unnamed_N"
                processed_headers.append(f"Unnamed_{unnamed_counter}")
                unnamed_counter += 1
            else:
                # Applica strip e aggiungi l'header non vuoto
                processed_headers.append(header.strip())
        
        # Ora gestiamo i duplicati
        header_counts = Counter()
        unique_headers = []
        
        for header in processed_headers:
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
    
    def export_df_to_excel(df, output_path=None, sheet_name="Avvisi", verify_columns=True):
        """
        Esporta un DataFrame in un file Excel dopo aver verificato la presenza di valori mancanti
        in colonne specifiche e altre verifiche di qualità dei dati.
        
        Args:
            df: Il DataFrame pandas da esportare
            output_path: Il percorso del file Excel di output. Se None, viene generato automaticamente
            sheet_name: Il nome del foglio Excel
            verify_columns: Se True, esegue verifiche sui dati prima dell'esportazione
            
        Returns:
            tuple: (successo, messaggio, percorso_file)
                - successo (bool): True se l'esportazione è andata a buon fine, False altrimenti
                - messaggio (str): Messaggio informativo o di errore
                - percorso_file (str): Percorso del file esportato in caso di successo
        """
        # Verifica che il DataFrame non sia vuoto
        if df is None or df.empty:
            return False, "Il DataFrame è vuoto o nullo.", None
        
        # Crea una copia per non modificare l'originale
        df_export = df.copy()
        
        # Se richiesto, esegue verifiche sui dati
        if verify_columns:
            # 1. Verifica che le colonne obbligatorie esistano
            required_columns = ["Avviso", "Data", "Sede tecnica", "St.sist."]
            missing_columns = [col for col in required_columns if col not in df_export.columns]
            
            if missing_columns:
                return False, f"Colonne obbligatorie mancanti: {', '.join(missing_columns)}", None
            
            # 2. Verifica valori mancanti nelle colonne obbligatorie
            missing_data = {}
            for col in required_columns:
                missing_count = df_export[col].isna().sum()
                if missing_count > 0:
                    missing_data[col] = missing_count
            
            if missing_data:
                missing_info = "\n".join([f"- {col}: {count} valori mancanti" for col, count in missing_data.items()])
                return False, f"Valori mancanti nelle colonne obbligatorie:\n{missing_info}", None
            
            # 3. Verifica formato numerico per la colonna Avviso
            if not pd.api.types.is_numeric_dtype(df_export["Avviso"]):
                try:
                    # Prova a convertire la colonna Avviso in numerico
                    df_export["Avviso"] = pd.to_numeric(df_export["Avviso"])
                    print("Colonna 'Avviso' convertita in formato numerico.")
                except Exception as e:
                    return False, f"La colonna 'Avviso' contiene valori non numerici: {str(e)}", None
            
            # 4. Verifica che i valori in St.sist. siano in un formato valido (es. solo certi valori)
            valid_st_sist = ["MELA", "MECO", "FCAN MECO", "MAPE", "MECO ORAT", "MELA ORAT"]
            invalid_st_sist = df_export[~df_export["St.sist."].isin(valid_st_sist)]["St.sist."].unique()
            
            if len(invalid_st_sist) > 0:
                print(f"Attenzione: Valori non standard nella colonna 'St.sist.': {', '.join(map(str, invalid_st_sist))}")
            
            # 5. Verifica che le date siano in un formato valido
            if not pd.api.types.is_datetime64_dtype(df_export["Data"]):
                try:
                    # Prova a convertire la colonna Data in datetime
                    df_export["Data"] = pd.to_datetime(df_export["Data"], dayfirst=True)
                    print("Colonna 'Data' convertita in formato datetime.")
                except Exception as e:
                    print(f"Attenzione: Non è stato possibile convertire la colonna 'Data' in formato datetime: {str(e)}")
                    # Continuiamo comunque con l'esportazione
            
            # 6. Verifica il formato della colonna Sede tecnica (pattern comune)
            # Es: MXW-MXPA-X4-18-GE-ES
            invalid_sede = []
            sede_pattern = re.compile(r'^[A-Z]{3}-[A-Z]{4}-[A-Z][0-9]-[0-9]{2}-[A-Z]{2}-[A-Z]{2}$')
            
            for idx, sede in enumerate(df_export["Sede tecnica"]):
                if sede and not sede_pattern.match(str(sede)) and not sede == "CLS-CLSB-07-03-IT":
                    invalid_sede.append((idx, sede))
            
            if invalid_sede:
                print(f"Attenzione: {len(invalid_sede)} valori nella colonna 'Sede tecnica' non seguono il pattern standard.")
                if len(invalid_sede) <= 5:  # Mostra solo i primi 5 esempi
                    for idx, sede in invalid_sede[:5]:
                        print(f"  Riga {idx+1}: '{sede}'")
        
        # Genera un nome di file se non fornito
        if output_path is None:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_path = f"avvisi_export_{timestamp}.xlsx"
        
        # Verifica che la directory di destinazione esista
        output_dir = os.path.dirname(output_path)
        if output_dir and not os.path.exists(output_dir):
            try:
                os.makedirs(output_dir)
                print(f"Creata directory: {output_dir}")
            except Exception as e:
                return False, f"Impossibile creare la directory di destinazione: {str(e)}", None
        
        # Esporta il DataFrame in Excel
        try:
            with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                df_export.to_excel(writer, sheet_name=sheet_name, index=False)
            
            # Verifica che il file sia stato creato correttamente
            if os.path.exists(output_path):
                file_size = os.path.getsize(output_path)
                row_count = len(df_export)
                return True, f"Esportazione completata con successo. File: {output_path}, Dimensione: {file_size/1024:.1f} KB, Righe: {row_count}", output_path
            else:
                return False, "Il file è stato creato ma non è possibile trovarlo nel percorso specificato.", None
        
        except Exception as e:
            return False, f"Errore durante l'esportazione in Excel: {str(e)}", None