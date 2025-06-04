def handle_extraction_result(status_code, result, tipo_estrazione, prefix=None):
    """
    Gestisce il risultato di un'estrazione IW39.
    
    Args:
        status_code: Codice di stato dell'estrazione
        result: Risultato dell'estrazione
        tipo_estrazione: Tipo di estrazione eseguita
        prefix: Prefisso opzionale per il logging
    
    Returns:
        bool: True se l'elaborazione è riuscita, False altrimenti
    """
    nonlocal totale_estratti
    
    # Funzione helper per creare il prefisso del log
    def get_log_prefix():
        return f"{prefix or ''} - {tipo_estrazione}"
    
    # Funzione helper per elaborare i dati e creare DataFrame
    def process_dataframe_creation():
        """Elabora il risultato e crea il DataFrame"""
        # Verifica e corregge il contenuto della clipboard
        success, fixed_content = self.fix_clipboard_table_content(result)
        if not success:
            self.log("Non è stato possibile correggere il contenuto della clipboard.", "error", True, True, 0)
            return None
        
        # Crea il DataFrame
        df = self.df_utils.clean_data(fixed_content)
        if df is None:
            return None
        
        # Verifica coerenza numero righe
        if status_code != len(df):
            self.log(f"Il DataFrame non ha lo stesso numero di righe della tabella SAP", "error", True, True, 0)
            return None
        
        # Aggiunge colonna tipo estrazione
        df['TipoEstrazione'] = tipo_estrazione
        return df
    
    # Funzione helper per salvare DataFrame
    def save_dataframe(df):
        """Salva il DataFrame nel dizionario iw39"""
        key = f"df_{tipo_estrazione}{f'_{prefix}' if prefix else ''}"
        iw39[key] = df
        self.log(f"DataFrame {key} creato con {len(df)} righe", "success", True, True, 0)
        return True
    
    # Gestione dei diversi codici di stato
    if status_code == -1:
        # Errore
        self.log(f"Fallita estrazione IW39 per {get_log_prefix()}", "error", True, True, 0)
        return False
    
    elif status_code == 0:
        # Nessun risultato
        self.log(f"Nessun dato trovato per {get_log_prefix()}", "info", True, True, 0)
        return True
    
    elif status_code == 1:
        # Singolo valore
        if tipo_estrazione != "ListaSingoli":
            # Aggiunge al set dei valori singoli
            self.log(f"Singolo valore trovato per {get_log_prefix()}", "info", True, True, 0)
            single_value_set.add(result)
            self.log(f"Aggiunto valore a single_value_set - numero elementi: {len(single_value_set)}", "info", True, True, 0)
            totale_estratti += status_code
            return True
        else:
            # Tratta come DataFrame anche se singolo valore
            self.log(f"Eseguita estrazione IW39 per {get_log_prefix()}", "success", True, True, 0)
            df = process_dataframe_creation()
            if df is None:
                self.log(f"DataFrame vuoto per {get_log_prefix()}", "error", True, True, 0)
                return False
            return save_dataframe(df)
    
    elif status_code > 1:
        # Lista con più elementi
        self.log(f"Eseguita estrazione IW39 per {get_log_prefix()}", "success", True, True, 0)
        df = process_dataframe_creation()
        if df is None:
            self.log(f"DataFrame vuoto per {get_log_prefix()}", "error", True, True, 0)
            return False
        
        # Incrementa contatore solo se non è "ListaSingoli"
        if tipo_estrazione != "ListaSingoli":
            totale_estratti += status_code
        
        return save_dataframe(df)
    
    else:
        # Valore non valido
        self.log(f"Valore di status_code non valido: {status_code}", "error", True, True, 0)
        return False