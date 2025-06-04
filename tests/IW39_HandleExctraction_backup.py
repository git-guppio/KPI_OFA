        def handle_extraction_result(status_code, result, tipo_estrazione, prefix=None):
            nonlocal totale_estratti
            if status_code == -1:  # codice_stato: -1=errore
                self.log(f"Fallita estrazione IW39 per {prefix or ''} - {tipo_estrazione}", "error", True, True, 0)
                return False
            elif status_code > 1:  # Successo con lista, status_code indica il numero di elementi presenti nella lista SAP
                self.log(f"Eseguita estrazione IW39 per {prefix or ''} - {tipo_estrazione}", "success", True, True, 0)
                # verifico la coerenza delle righe nel risultato (presenza del carattere #)
                success, fixed_content = self.fix_clipboard_table_content(result)
                if not success:
                    self.log(f"Non è stato possibile correggere il contenuto della clipboard.", "error", True, True, 0)
                    return False                    
                # La chiave sarà 'df_Creazione', 'df_Modifica', ecc.
                key = f"df_{tipo_estrazione}{f'_{prefix}' if prefix else ''}"
                df = self.df_utils.clean_data(fixed_content)
                if df is None:
                    self.log(f"DataFrame vuoto per {key}", "error", True, True, 0)
                    return False
                # Verifica se il DataFrame contenga lo stesso numero di righe della tabella SAP da cui sono stati estratti i dati
                if (status_code != len(df)):
                    self.log(f"Il DataFrame {key} non ha lo stesso numero di righe della tabella SAP", "error", True, True, 0)
                    return False              
                # aggiungo la colonna con la tipologia di estrazione per tenere traccia
                df['TipoEstrazione'] = tipo_estrazione
                iw39[key] = df
                self.log(f"DataFrame {key} creato con {len(df)} righe", "success", True, True, 0)
                # Incremento il contatore totale degli estratti solo per le estrazioni != da "ListaSingoli"
                if(tipo_estrazione!="ListaSingoli"):
                    totale_estratti += status_code                
                return True
            elif status_code == 1:  # Singolo valore
                if (tipo_estrazione != "ListaSingoli"):
                    # Se il tipo di estrazione non è "ListaSingoli", aggiungo il valore al set   
                    self.log(f"Singolo valore trovato per {prefix or ''} - {tipo_estrazione}", "info", True, True, 0)
                    self.log(f"Aggiunto valore a single_value_set - numero elementi: {len(single_value_set)}", "info", True, True, 0)
                    single_value_set.add(result)  # set.add() aggiunge solo se non esiste già
                    # Incremento il contatore totale degli estratti, il caso == 1 è per tutti i tipi di estrazione
                    totale_estratti += status_code               
                    return True
                else:
                    self.log(f"Eseguita estrazione IW39 per {prefix or ''} - {tipo_estrazione}", "success", True, True, 0)
                    # verifico la coerenza delle righe nel risultato (presenza del carattere #)
                    success, fixed_content = self.fix_clipboard_table_content(result)
                    if not success:
                        self.log(f"Non è stato possibile correggere il contenuto della clipboard.", "error", True, True, 0)
                        return False                    
                    # La chiave sarà 'df_Creazione', 'df_Modifica', ecc.
                    key = f"df_{tipo_estrazione}{f'_{prefix}' if prefix else ''}"
                    df = self.df_utils.clean_data(fixed_content)
                    if df is None:
                        self.log(f"DataFrame vuoto per {key}", "error", True, True, 0)
                        return False
                    # Verifica se il DataFrame contenga lo stesso numero di righe della tabella SAP da cui sono stati estratti i dati
                    if (status_code != len(df)):
                        self.log(f"Il DataFrame {key} non ha lo stesso numero di righe della tabella SAP", "error", True, True, 0)
                        return False              
                    # aggiungo la colonna con la tipologia di estrazione per tenere traccia
                    df['TipoEstrazione'] = tipo_estrazione
                    iw39[key] = df
                    self.log(f"DataFrame {key} creato con {len(df)} righe", "success", True, True, 0)                    

            elif status_code == 0:  # Nessun risultato
                self.log(f"Nessun dato trovato per {prefix or ''} - {tipo_estrazione}", "info", True, True, 0)
                return True
            else:
                self.log(f"Valore di status_code non valido: {status_code}", "error", True, True, 0)
                return False