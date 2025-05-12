import time
import pyperclip
import win32con
import pandas as pd

from collections import Counter
from typing import List, Dict, Optional
import Config.constants as constants
import os
from typing import Dict, Any, Optional
import logging
from utils.decorators import error_logger
import DF_Tools

# Logger specifico per questo modulo
logger = logging.getLogger("DataFrameTools")


class SAPDataExtractor:
    """
    Classe per eseguire estrazioni dati da SAP utilizzando una sessione esistente
    """

    def __init__(self, session):
        """
        Inizializza la classe con una sessione SAP attiva
        
        Args:
            session: Oggetto sessione SAP attiva
        """
        self.session = session
        
        self.tipo_estrazioni = ["Creazione", "Modifica", "Lista"]
        self.df_utils = DF_Tools.DataFrameTools()

    def copy_values_for_sap_selection(self, values):
        """
        Copia valori formattati per un campo di selezione multipla SAP.
        
        Args:
            values: Lista o set di valori da copiare
        """
        try:
            if values.empty:
                logger.warning("Nessun valore da copiare")
                return False
            
            # Assicura che sia una lista
            if isinstance(values, set):
                values = list(values)
            
            # Formatta i valori uno per riga (formato accettato da SAP per selezioni multiple)
            text = '\r\n'.join(str(item) for item in values)
            
            # Copia nella clipboard
            pyperclip.copy(text)
            time.sleep(0.1)
            # Alternativa con win32clipboard
            """
            win32clipboard.OpenClipboard()
            win32clipboard.EmptyClipboard()
            win32clipboard.SetClipboardText(text, win32con.CF_TEXT)
            win32clipboard.CloseClipboard() 
            """
            
            logger.info(f"Copiati {len(values)} valori nella clipboard per SAP")
            return True
        except Exception as e:
            logger.error(f"Errore durante la copia nella clipboard: {str(e)}")
            return False        

    def extract_IW29_old(self, dataInizio, dataFine, tech_config, lista_AdM ) -> tuple[bool, pd.DataFrame | None]:
        # Crea un dizionario vuoto per memorizzare i DataFrame
        iw29 = {}
        # Inizializzazione di un set per memorizzare dati univoci
        single_value_set = set()
        # Itera attraverso le estrazioni
        for tipo_estrazione in self.tipo_estrazioni:
            # Estrazione dati per ogni tecnologia configurata
            for tech, prefixes in tech_config.items():
                if not prefixes:
                    logger.warning(f"Nessun prefisso configurato per {tech}, skip")
                    continue
                for prefix in prefixes:
                    # Verifica se il prefisso è vuoto o contiene solo spazi
                    if not prefix.strip():
                        logger.warning(f"Prefisso vuoto per {tech}, skip")
                        continue
                    logger.info(f"Estrazione IW29_single per {tech} - {prefix} - {tipo_estrazione}")
                    # Esegui l'estrazione e ottieni una lista di dizionari
                    str_dataInizio = dataInizio.toString("dd.MM.yyyy")  # Formato gg.mm.aaaa
                    str_dataFine = dataFine.toString("dd.MM.yyyy")  # Formato gg.mm.aaaa
                    status_code, result = self.extract_IW29_single(str_dataInizio, str_dataFine, tipo_estrazione, lista_AdM, prefix)
                    # Verifico l'esito dell'estrazione
                    if status_code == 0: # codice_stato: 0=errore, 1=successo con lista, 2=singolo valore, 3=nessun risultato
                        logger.error(f"Fallita estrazione IW29_single per {prefix} - {tipo_estrazione}")
                    elif status_code == 1:  # Successo con lista
                        logger.info(f"Eseguita estrazione IW29_single per {prefix} - {tipo_estrazione}")
                        # La chiave sarà 'df_Creazione', 'df_Modifica', ecc.
                        key = f"df_{tipo_estrazione}{f'_{prefix}' if prefix else ''}"
                        df = self.df_utils.clean_data(result)
                        if df is None:
                            logger.error(f"DataFrame vuoto per {key}")
                            return False
                        # aggiungo la colonna con la tipologia di estrazione per tenere traccia
                        df['TipoEstrazione'] = tipo_estrazione
                        iw29[key] = df
                        logger.info(f"DataFrame {key} creato con {len(df)} righe")
                    elif status_code == 2:  # Singolo valore
                        logger.info(f"Singolo valore trovato per {prefix} - {tipo_estrazione}")
                        single_value_set.add(result)  # set.add() aggiunge solo se non esiste già
                        # Gestisci il singolo valore come necessario
                    elif status_code == 3:
                        logger.info(f"Nessun dato trovato per {prefix} - {tipo_estrazione}")
        #verifico al termine dei cicli se la lista contenente i valori singoli è vuota
        if len(single_value_set) == 0:
            logger.info("Nessun valore singolo trovato")
        # Se rilevo un solo valore allora ne prelevo uno fra quelli già presente nel df per eseguire la estrazione di un lista
        # il valore verrà poi eliminato come duplicato
        elif len(single_value_set) == 1:
            logger.info(f"Un solo valore singolo trovato: {self.single_value}")
            # prelevo un Avviso dalla collezione dei df
            avviso_singolo = None
            for key, df in iw29.items():
                if not df.empty and "Avviso" in df.columns:
                    # Prendi il primo avviso disponibile e interrompi il ciclo
                    avviso_singolo = df["Avviso"].iloc[0]
                    break
            single_value_set.add(avviso_singolo)
        # verifico che la lista dei valori singoli contenga più di un valore    
        if len(single_value_set) > 1:
            # Converto il set in una lista
            list_value = list(single_value_set)
            logger.info(f"Valori singoli trovati: {list_value}")
            # ripeto l'estrazione per i valori risultanti
            tipo_estrazione = "ListaSingoli"
            status_code, result = self.extract_IW29_single(str_dataInizio, str_dataFine, tipo_estrazione, list_value, None)
            # Verifico l'esito dell'estrazione
            if status_code == 0: # codice_stato: 0=errore, 1=successo con lista, 2=singolo valore, 3=nessun risultato
                logger.error(f"Fallita estrazione IW29_single per lista valori singoli")
                return False
            elif status_code == 1:  # Successo con lista
                logger.info(f"Eseguita estrazione IW29_single per lista valori singoli")
                # La chiave sarà 'df_Creazione', 'df_Modifica', ecc.
                key = f"df_{tipo_estrazione}"
                df = self.df_utils.clean_data(result)
                # aggiungo la colonna con la tipologia di estrazione per tenere traccia
                df['TipoEstrazione'] = tipo_estrazione
                iw29[key] = df
                logger.info(f"DataFrame {key} creato con {len(df)} righe")
            elif status_code == 2:  # Singolo valore
                logger.info(f"Singolo valore trovato per {tipo_estrazione}")
                return False
            elif status_code == 3:
                logger.info(f"Nessun dato trovato per {tipo_estrazione}")
                return False     



        # Unisci i DataFrame in un unico DataFrame finale
        if iw29:
            df_finale = pd.concat(iw29.values(), ignore_index=True)
            return True, df_finale

# ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    def extract_IW29(self, dataInizio, dataFine, tech_config, lista_AdM) -> tuple[bool, pd.DataFrame | None]:
        # Crea un dizionario vuoto per memorizzare i DataFrame
        iw29 = {}
        # Inizializzazione di un set per memorizzare dati univoci
        single_value_set = set()
        
        # Converti le date in stringhe per riutilizzarle in tutto il metodo
        str_dataInizio = dataInizio.toString("dd.MM.yyyy")  # Formato gg.mm.aaaa
        str_dataFine = dataFine.toString("dd.MM.yyyy")  # Formato gg.mm.aaaa
        
        # Funzione interna per gestire il risultato dell'estrazione
        def handle_extraction_result(status_code, result, tipo_estrazione, prefix=None):
            if status_code == 0:  # codice_stato: 0=errore
                logger.error(f"Fallita estrazione IW29_single per {prefix or ''} - {tipo_estrazione}")
                return False
            elif status_code == 1:  # Successo con lista
                logger.info(f"Eseguita estrazione IW29_single per {prefix or ''} - {tipo_estrazione}")
                # verifico la coerenza delle righe nel risultato (pesenza del carattere #)
                success, fixed_content = self.fix_clipboard_table_content(result)
                if not success:
                    logger.error(f"Non è stato possibile correggere il contenuto della clipboard.")
                    return False                    
                # La chiave sarà 'df_Creazione', 'df_Modifica', ecc.
                key = f"df_{tipo_estrazione}{f'_{prefix}' if prefix else ''}"
                df = self.df_utils.clean_data(fixed_content)
                if df is None:
                    logger.error(f"DataFrame vuoto per {key}")
                    return False
                # aggiungo la colonna con la tipologia di estrazione per tenere traccia
                df['TipoEstrazione'] = tipo_estrazione
                iw29[key] = df
                logger.info(f"DataFrame {key} creato con {len(df)} righe")
                return True
            elif status_code == 2:  # Singolo valore
                logger.info(f"Singolo valore trovato per {prefix or ''} - {tipo_estrazione}")
                single_value_set.add(result)  # set.add() aggiunge solo se non esiste già
                return True
            elif status_code == 3:  # Nessun risultato
                logger.info(f"Nessun dato trovato per {prefix or ''} - {tipo_estrazione}")
                return True
            return False
        
        # Gestisci prima l'estrazione di tipo "Lista", che deve essere eseguita una sola volta
        if "Lista" in self.tipo_estrazioni:
            logger.info(f"Estrazione IW29_single per tipo 'Lista' con tutti i valori")
            # Esegui l'estrazione per il tipo "Lista" senza iterare su tech e prefissi
            status_code, result = self.extract_IW29_single(str_dataInizio, str_dataFine, "Lista", lista_AdM, "")
            
            # Utilizzo della funzione interna per gestire il risultato
            if not handle_extraction_result(status_code, result, "Lista"):
                return False, None
        
        # Itera attraverso le estrazioni, escludendo "Lista" che è già stata gestita
        for tipo_estrazione in [t for t in self.tipo_estrazioni if t != "Lista"]:
            # Estrazione dati per ogni tecnologia configurata
            for tech, prefixes in tech_config.items():
                if not prefixes:
                    logger.warning(f"Nessun prefisso configurato per {tech}, skip")
                    continue
                for prefix in prefixes:
                    # Verifica se il prefisso è vuoto o contiene solo spazi
                    if not prefix.strip():
                        logger.warning(f"Prefisso vuoto per {tech}, skip")
                        continue
                    logger.info(f"Estrazione IW29_single per {tech} - {prefix} - {tipo_estrazione}")
                    # Esegui l'estrazione e ottieni una lista di dizionari
                    status_code, result = self.extract_IW29_single(str_dataInizio, str_dataFine, tipo_estrazione, lista_AdM, prefix)
                    
                    # Utilizzo della funzione interna per gestire il risultato
                    if not handle_extraction_result(status_code, result, tipo_estrazione, prefix):
                        return False, None
        
        # verifico al termine dei cicli se la lista contenente i valori singoli è vuota
        if len(single_value_set) == 0:
            logger.info("Nessun valore singolo trovato")
        # Se rilevo un solo valore allora ne prelevo uno fra quelli già presente nel df per eseguire la estrazione di un lista
        # il valore verrà poi eliminato come duplicato
        elif len(single_value_set) == 1:
            logger.info(f"Un solo valore singolo trovato: {next(iter(single_value_set))}")
            # prelevo un Avviso dalla collezione dei df
            avviso_singolo = None
            for key, df in iw29.items():
                if not df.empty and "Avviso" in df.columns:
                    # Prendi il primo avviso disponibile e interrompi il ciclo
                    avviso_singolo = df["Avviso"].iloc[0]
                    break
            if avviso_singolo:
                single_value_set.add(avviso_singolo)
        
        # verifico che la lista dei valori singoli contenga più di un valore    
        if len(single_value_set) > 1:
            # Converto il set in una lista
            list_value = list(single_value_set)
            logger.info(f"Valori singoli trovati: {list_value}")
            # ripeto l'estrazione per i valori risultanti
            status_code, result = self.extract_IW29_single(str_dataInizio, str_dataFine, "ListaSingoli", list_value, None)
            
            # Utilizzo della funzione interna per gestire il risultato
            if not handle_extraction_result(status_code, result, "ListaSingoli"):
                return False, None
        
        # Se il dizionario di DataFrame è vuoto, restituisci False
        if not iw29:
            logger.error("Nessun DataFrame creato")
            return False, None
        
        # Concatena tutti i DataFrame in un unico DataFrame
        result_df = pd.concat(iw29.values(), ignore_index=True) if iw29 else None
        # Rimuovi le righe duplicate
        result_df = result_df.drop_duplicates()
        return True, result_df
# ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    def extract_IW29_single(self, dataInizio, dataFine, tipo_estrazione, lista_AdM, prefix=None) -> tuple[int, str | None]:
        """
        Estrae dati relativi agli avvisi di manutenzione utilizzando la transazione IW29

        Args:
            dataInizio (str): Data di inizio nel formato 'dd.MM.yyyy'
            dataFine (str): Data di fine nel formato 'dd.MM.yyyy'
            tipo_estrazione (str): Tipo di estrazione (Creazione, Modifica, Lista)
            prefix (str, optional): Prefisso che indica l'impianto da considerare.        
            
        Returns:
            tuple: (codice_stato, dati) dove:
                - codice_stato: 0=errore, 1=successo con lista, 2=singolo valore, 3=nessun risultato
                - dati: lista di dizionari o messaggio di errore

        Raises:
            DataReturnedError: Errori durante l'estrazione dei dati
            ConnectionError: Se ci sono problemi di connessione con SAP
        """       
        try:
            # Svuota la clipboard prima dell'estrazione
            # Svuota la clipboard copiando una stringa vuota
            pyperclip.copy("")
            time.sleep(0.1)
            # Alternativa con win32clipboard
            """ 
            win32clipboard.OpenClipboard()
            win32clipboard.EmptyClipboard()
            win32clipboard.CloseClipboard()
             """
            #self.session.findById("wnd[0]").resizeWorkingPane(173, 49, False)
            self.session.findById("wnd[0]/tbar[0]/okcd").text = "/nIW29"
            self.session.findById("wnd[0]").sendVKey(0)
        # tutti gli stati
            self.session.findById("wnd[0]/usr/chkDY_OFN").selected = True
            self.session.findById("wnd[0]/usr/chkDY_IAR").selected = True
            self.session.findById("wnd[0]/usr/chkDY_RST").selected = True
            self.session.findById("wnd[0]/usr/chkDY_MAB").selected = True
        # elimino data avviso
            self.session.findById("wnd[0]/usr/ctxtDATUV").text = ""
            self.session.findById("wnd[0]/usr/ctxtDATUB").text = ""
            #self.session.findById("wnd[0]/usr/ctxtDATUB").setFocus()
            #self.session.findById("wnd[0]/usr/ctxtDATUB").caretPosition = 0
        # tipologia avvisi            
            self.session.findById("wnd[0]/usr/btn%_QMART_%_APP_%-VALU_PUSH").press()
            time.sleep(0.25)
            self.session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = "Z1"
            self.session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "Z2"
            self.session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").text = "Z3"
            self.session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,3]").text = "Z4"
            self.session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,4]").text = "Z5"
            #self.session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLA++++LDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,4]").setFocus()
            self.session.findById("wnd[1]/tbar[0]/btn[8]").press()
            # rimuovo il valore dal data avviso
            self.session.findById("wnd[0]/usr/ctxtDATUB").text = ""
            self.session.findById("wnd[0]/usr/ctxtDATUV").text = ""


            if (tipo_estrazione == "Creazione"):
                if(prefix == None):
                    raise ValueError("Atteso un prefisso per l'estrazione di tipo Creazione")
                # Sede tecnica    WSB
                self.session.findById("wnd[0]/usr/ctxtSTRNO-LOW").text = f"{prefix}-++++*"                
                # Data Creazione - inizio e fine mese
                self.session.findById("wnd[0]/usr/ctxtERDAT-LOW").text = dataInizio 
                self.session.findById("wnd[0]/usr/ctxtERDAT-HIGH").text =  dataFine
                # Data Modifica - inizio e fine mese
                self.session.findById("wnd[0]/usr/ctxtAEDAT-LOW").text = ""
                self.session.findById("wnd[0]/usr/ctxtAEDAT-HIGH").text = ""
            elif (tipo_estrazione == "Modifica"):
                # Sede tecnica    WSB
                self.session.findById("wnd[0]/usr/ctxtSTRNO-LOW").text = f"{prefix}-++++*"
                # Data Creazione - inizio e fine mese
                self.session.findById("wnd[0]/usr/ctxtERDAT-LOW").text = "" 
                self.session.findById("wnd[0]/usr/ctxtERDAT-HIGH").text =  ""
                # Data Modifica - inizio e fine mese
                self.session.findById("wnd[0]/usr/ctxtAEDAT-LOW").text = dataInizio
                self.session.findById("wnd[0]/usr/ctxtAEDAT-HIGH").text = dataFine
            elif (tipo_estrazione == "Lista") or (tipo_estrazione == "ListaSingoli"):
                # Copia i valori della lista nella clipboard per l'uso in SAP
                if len(lista_AdM) == 0:
                    raise ValueError("Nessun valore nella lista AdM da copiare nella clipboard")
                else:
                    if not(self.copy_values_for_sap_selection(lista_AdM)):
                        raise ValueError("Errore durante la copia dei valori nella clipboard")
                logger.info("Valori singoli copiati nella clipboard per SAP")

                self.session.findById("wnd[0]/usr/ctxtSTRNO-LOW").text = ""                
                self.session.findById("wnd[0]/usr/ctxtDATUV").text = ""
                self.session.findById("wnd[0]/usr/ctxtDATUB").text = ""
                self.session.findById("wnd[0]/usr/ctxtAEDAT-LOW").text = ""
                self.session.findById("wnd[0]/usr/ctxtAEDAT-HIGH").text = ""
                # incollo i valori nella clipboard
                self.session.findById("wnd[0]/usr/btn%_QMNUM_%_APP_%-VALU_PUSH").press()
                time.sleep(0.25)
                self.session.findById("wnd[1]/tbar[0]/btn[24]").press()
                time.sleep(0.25)
                self.session.findById("wnd[1]/tbar[0]/btn[8]").press()
                time.sleep(0.25)
            else:
                raise ValueError(f"Tipo di estrazione non valido: {tipo_estrazione}")
            #self.session.findById("wnd[0]").sendVKey(0)

        # Layout

            self.session.findById("wnd[0]/usr/ctxtVARIANT").text = "/KPIOFANO2"
            #self.session.findById("wnd[0]/usr/ctxtVARIANT").setFocus()
            #self.session.findById("wnd[0]/usr/ctxtVARIANT").caretPosition = 10
            #self.session.findById("wnd[0]").sendVKey(0)

        # Esegui

            self.session.findById("wnd[0]/tbar[1]/btn[8]").press()
            # Attendi che SAP sia pronto
            if not self.wait_for_sap(30):
                msg = "Timeout durante l'esecuzione della transazione SAP IW29"
                logger.warning(msg)
                return 0, msg # codice_stato: 0=errore, 1=successo con lista, 2=singolo valore, 3=nessun risultato
            time.sleep(0.5)
            # Verifico che siano stati estratti dei dati
            if self.session.findById("wnd[0]/sbar").text == "Non sono stati selezionati oggetti":
                msg =  "Nessun dato trovato"
                logger.warning(msg)
                return 3, msg # codice_stato: 0=errore, 1=successo con lista, 2=singolo valore, 3=nessun risultato
            # Verifico se è stato estratto un solo valore
            if self.session.findById("wnd[0]").text == "Visualizzare avviso PM: Segnalazione guasto":
                msg = "Un solo valore trovato"
                logger.info(msg)
                AdM = self.session.findById("wnd[0]/usr/subSCREEN_1:SAPLIQS0:1050/txtVIQMEL-QMNUM").text
                return 2, AdM # codice_stato: 0=errore, 1=successo con lista, 2=singolo valore, 3=nessun risultato            
            if (self.session.findById("wnd[0]").text == "Visualizzare avvisi: lista avvisi"):      # Titolo della finestra
                # Salvo i dati nella clipboard
                self.session.findById("wnd[0]/mbar/menu[0]/menu[11]/menu[2]").select()
                time.sleep(0.25)
                self.session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[4,0]").select()
                time.sleep(0.25)
                self.session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[4,0]").setFocus()
                time.sleep(0.25)
                # Svuota la clipboard per rilevare la corretta scrittura di nuovi dati
                try:
                    pyperclip.copy("")
                    print("Clipboard svuotata")
                except Exception as e:
                    print(f"Errore durante lo svuotamento della clipboard: {str(e)}")
                    return 0, e
                self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
                # Attendi che SAP sia pronto
                if not self.wait_for_sap(30):
                    msg = "Timeout durante il caricamento dei dati nella clipboard"
                    logger.error(msg)
                    return 0, msg # codice_stato: 0=errore, 1=successo con lista, 2=singolo valore, 3=nessun risultato
                time.sleep(1)
                # Attendi che la clipboard sia riempita
                if not self.wait_for_write_clipboard_data(30):
                    # Gestisci il caso in cui non sono stati trovati dati
                    msg = "Errore durante il caricamento dei dati nella clipboard"
                    logger.error(msg)
                    return 0, msg # codice_stato: 0=errore, 1=successo con lista, 2=singolo valore, 3=nessun risultato
                # Leggo il contenuto della clipboard
                data = pyperclip.paste() # Il tipo di dato è una stringa
                logger.info("Dati trovati nella clipboard")
                return 1, data
            # Se arriviamo qui, la condizione della finestra non è stata riconosciuta
            msg = "Stato SAP non riconosciuto"
            logger.warning(msg)
            return 0, msg
            
        except Exception as e:
            # Gestione generale degli errori
            logger.error(f"Errore nell'estrazione IW29_single: {str(e)}")
            return 0, str(e)

# ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

    def wait_for_sap(self, timeout: int = 30):  # timeout in secondi
        """
        Attende che SAP finisca le operazioni in corso
        
        Args:
            timeout: Tempo massimo di attesa in secondi
        
        Returns:
            bool: True se SAP è diventato disponibile, False se è scaduto il timeout
        """
        start_time = time.time()
        
        try:
            while self.session.Busy:
                # Verifica timeout
                if time.time() - start_time > timeout:
                    print(f"Timeout dopo {timeout} secondi di attesa")
                    return False
                    
                time.sleep(0.5)
                print("SAP is busy")
            
            return True
            
        except Exception as e:
            print(f"Errore durante l'attesa: {str(e)}")
            return False
            
    def wait_for_write_clipboard_data(self, timeout: int = 30) -> bool:
        """
        Attende che la clipboard contenga dei dati
        
        Args:
            timeout: Tempo massimo di attesa in secondi
            
        Returns:
            bool: True se sono stati trovati dati, False se è scaduto il timeout
        """
        
        start_time = time.time()
        last_print_time = 0  # Per limitare i messaggi di log
        print_interval = 2   # Intervallo in secondi tra i messaggi di log
        
        while True:
            current_time = time.time()
            
            # Verifica timeout
            if current_time - start_time > timeout:
                print(f"Timeout: nessun dato trovato nella clipboard dopo {timeout} secondi")
                return False
            
            try:
                # Controlla il contenuto della clipboard
                data = pyperclip.paste()
                
                # Verifica se ci sono dati nella clipboard
                if data and data.strip():
                    print("Dati trovati nella clipboard")
                    return True
                
                # Stampa il messaggio di attesa solo ogni print_interval secondi
                if current_time - last_print_time >= print_interval:
                    print("In attesa dei dati nella clipboard...")
                    last_print_time = current_time
                
                # Aspetta prima del prossimo controllo
                time.sleep(0.1)  # Ridotto il tempo di attesa per una risposta più veloce
                
            except Exception as e:
                print(f"Errore durante il controllo della clipboard: {str(e)}")
                time.sleep(0.5)  # Attesa più lunga in caso di errore
                continue 

# ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    def fix_clipboard_table_content(self, result: str) -> tuple[bool, str]:
        """
        Elabora il contenuto della tabella proveniente dalla clipboard.
        A partire dalla quarta riga, verifica che ogni riga inizi con il carattere '|'.
        Se una riga non inizia con '|', unisce il suo contenuto alla riga precedente.
        
        Args:
            result: Il contenuto della clipboard da elaborare
            
        Returns:
            tuple: (successo, contenuto_elaborato)
                - successo (bool): True se l'elaborazione è andata a buon fine, False altrimenti
                - contenuto_elaborato (str): Il contenuto elaborato se successo, altrimenti il contenuto originale
        """
        try:
            # Dividi il testo in righe
            lines = result.split('\n')
            
            # Verifica se ci sono almeno 4 righe
            if len(lines) < 4:
                print("Il contenuto ha meno di 4 righe. Impossibile procedere.")
                return False, result
            
            # Preserva le prime 3 righe
            processed_lines = lines[:3]
            
            # Contatore delle sostituzioni
            replacements_count = 0
            
            # Elabora a partire dalla quarta riga
            i = 3
            while i < len(lines):
                current_line = lines[i]
                
                # Verifica se la riga inizia con '|'
                if not current_line.startswith('|'):
                    # Verifica se c'è una riga precedente da modificare
                    if i > 0 and len(processed_lines) > 0:
                        prev_line = processed_lines[-1]
                        
                        # Trova l'ultima occorrenza di '|' nella riga precedente
                        last_pipe_index = prev_line.rfind('|')
                        
                        if last_pipe_index != -1:
                            # Estrai la parte iniziale fino all'ultimo '|'
                            prefix = prev_line[:last_pipe_index + 1]
                            
                            # Unisci con la riga corrente
                            merged_line = prefix + " " + current_line.strip()
                            
                            # Sostituisci la riga precedente
                            processed_lines[-1] = merged_line
                            
                            replacements_count += 1
                        else:
                            # Se non troviamo '|' nella riga precedente, semplicemente aggiungi la riga corrente
                            processed_lines.append(current_line)
                    else:
                        # Se non c'è una riga precedente, aggiungi semplicemente la riga
                        processed_lines.append(current_line)
                else:
                    # La riga inizia con '|', aggiungila normalmente
                    processed_lines.append(current_line)
                
                # Passa alla riga successiva
                i += 1
            
            # Ricostruisci il testo elaborato
            processed_result = '\n'.join(processed_lines)
            
            # Stampa il numero di sostituzioni
            print(f"Elaborazione completata. Sono state eseguite {replacements_count} sostituzioni.")
            
            # Restituisci il risultato dell'elaborazione
            return True, processed_result
            
        except Exception as e:
            print(f"Errore durante l'elaborazione del contenuto: {str(e)}")
            return False, result

""" 
def main():

#   Esempio di utilizzo combinato delle classi SAPGuiConnection e SAPDataExtractor

    from sap_gui_connection import SAPGuiConnection  # Importa la classe creata precedentemente
    
    try:
        # Utilizzo con context manager per la connessione
        with SAPGuiConnection() as sap:
            if sap.is_connected():
                # Ottieni la sessione
                session = sap.get_session()
                if session:
                    # Crea l'estrattore
                    extractor = SAPDataExtractor(session)
                    
                    # Estrai i materiali
                    print("Estrazione materiali...")
                    materials = extractor.extract_materials("1000")  # Plant code esempio
                    for material in materials:
                        print(f"Materiale: {material['material_code']} - {material['description']}")
                    
                    # Estrai gli ordini
                    print("\nEstrazione ordini...")
                    orders = extractor.extract_orders("20240101", "20240131")
                    for order in orders:
                        print(f"Ordine: {order['order_number']} - Cliente: {order['customer']}")
    
    except Exception as e:
        print(f"Errore generale: {str(e)}")


if __name__ == "__main__":
    main()
"""