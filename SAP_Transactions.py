import time
import win32clipboard
import pandas as pd

from collections import Counter
from typing import List, Dict, Optional
import Config.constants as constants
import os
from typing import Dict, Any, Optional
import logging
from utils.decorators import error_logger

# Logger specifico per questo modulo
logger = logging.getLogger("DataFrameTools")

class NoDataReturnedError(Exception):
    """
    Eccezione sollevata quando una transazione SAP non restituisce alcun dato.
    
    Questa eccezione viene generata quando una transazione SAP viene eseguita correttamente
    dal punto di vista tecnico, ma non produce risultati. Usata per distinguere l'assenza
    di dati da un errore di esecuzione.
    
    Attributes:
        message (str): Spiegazione dell'errore
    """
    pass

class SAPTransactionEmptyResultError(Exception):
    """
    Eccezione sollevata quando una transazione SAP restituisce un risultato vuoto.
    
    Questa eccezione viene generata quando ci si aspetta che una transazione SAP
    produca dei record, ma il risultato è un set vuoto. Utile per gestire casi in cui
    l'assenza di dati potrebbe richiedere una logica di fallback.
    
    Attributes:
        message (str): Spiegazione dell'errore
        transaction_code (str, optional): Il codice della transazione SAP che ha prodotto il risultato vuoto
    """
    pass

class SAPQueryResultEmptyError(Exception):
    """
    Eccezione sollevata quando una query SAP non produce risultati.
    
    Questa eccezione è specifica per le operazioni di interrogazione/query su SAP
    che non generano record. Permette di distinguere le query vuote da altri tipi
    di operazioni SAP senza risultati.
    
    Attributes:
        message (str): Spiegazione dell'errore
        query_id (str, optional): Identificativo della query eseguita
        parameters (dict, optional): Parametri utilizzati nella query
    """
    pass


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
    
    def extract_IW29(self, dataInizio, dataFine, tech_config) -> pd.DataFrame | None:
        # Crea un dizionario vuoto per memorizzare i DataFrame
        iw29 = {}
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
                    # Esegui l'estrazione e ottieni una lista di dizionari
                    str_dataInizio = dataInizio.toString("dd.MM.yyyy")  # Formato gg.mm.aaaa
                    str_dataFine = dataFine.toString("dd.MM.yyyy")  # Formato gg.mm.aaaa
                    success, iw29_lista = self.extract_IW29_single(str_dataInizio, str_dataFine, tipo_estrazione, prefix)
                    # Se l'estrazione è riuscita, elabora i dati e memorizzali in un DataFrame nel dizioanrio iw29
                    if success:
                        # La chiave sarà 'df_Creazione', 'df_Modifica', ecc.
                        key = f"df_{tipo_estrazione}"
                        df = self.df_utils.clean_data(iw29_lista)
                        iw29[key] = df
                        logger.info(f"DataFrame {key} creato con {len(df)} righe")
                    else:
                        logger.warning(f"Estrazione {tipo_estrazione} non riuscita per {tech}: {prefix}")

    def extract_IW29_single(self, dataInizio, dataFine, tipo_estrazione, prefix) -> List[Dict] | None:
        """
        Estrae dati relativi agli avvisi di manutenzione utilizzando la transazione IW29
            
        Returns:
            Una lista di dizionari contenente i dati estratti dalla clipboard

        Raises:
            DataReturnedError: Errori durante l'estrazione dei dati
            ConnectionError: Se ci sono problemi di connessione con SAP
        """        
        try:
            # Svuota la clipboard prima dell'estrazione
            win32clipboard.OpenClipboard()
            win32clipboard.EmptyClipboard()
            win32clipboard.CloseClipboard()
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
        # Sede tecnica    WSB
            self.session.findById("wnd[0]/usr/ctxtSTRNO-LOW").text = f"{prefix}-++++*"
        # Data Creazione - inizio e fine mese (da poter inserire e modificare
            if (tipo_estrazione == "Creazione"):
                self.session.findById("wnd[0]/usr/ctxtDATUV").text = dataInizio
                self.session.findById("wnd[0]/usr/ctxtDATUB").text = dataFine
                self.session.findById("wnd[0]/usr/ctxtAEDAT-LOW").text = ""
                self.session.findById("wnd[0]/usr/ctxtAEDAT-HIGH").text = ""
            elif (tipo_estrazione == "Modifica"):
                self.session.findById("wnd[0]/usr/ctxtDATUV").text = ""
                self.session.findById("wnd[0]/usr/ctxtDATUB").text = ""
                self.session.findById("wnd[0]/usr/ctxtAEDAT-LOW").text = dataInizio
                self.session.findById("wnd[0]/usr/ctxtAEDAT-HIGH").text = dataFine
            elif (tipo_estrazione == "Lista"):
                self.session.findById("wnd[0]/usr/ctxtDATUV").text = ""
                self.session.findById("wnd[0]/usr/ctxtDATUB").text = ""
                self.session.findById("wnd[0]/usr/ctxtAEDAT-LOW").text = ""
                self.session.findById("wnd[0]/usr/ctxtAEDAT-HIGH").text = ""
                self.session.findById("wnd[0]/usr/btn%_QMNUM_%_APP_%-VALU_PUSH").press()
                time.sleep(0.25)
                self.session.findById("wnd[1]/tbar[0]/btn[24]").press()
                time.sleep(0.25)
                self.session.findById("wnd[1]/tbar[0]/btn[8]").press()
                time.sleep(0.25)
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
                print(f"Timeout durante l'esecuzione della transazione")
                return False
            time.sleep(0.5)
            # Verifico che siano stati estratti dei dati
            if self.session.findById("wnd[0]/sbar").text == "Non sono stati selezionati oggetti":
                print("Nessun dato trovato")
                return False
            if (self.session.findById("wnd[0]").text == "Visualizzare avvisi: lista avvisi"):      # Titolo della finestra
                # Salvo i dati nella clipboard
                self.session.findById("wnd[0]/mbar/menu[0]/menu[11]/menu[2]").select()
                time.sleep(0.25)
                self.session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[4,0]").select()
                time.sleep(0.25)
                self.session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[4,0]").setFocus()
                time.sleep(0.25)
                self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
                # Attendi che SAP sia pronto
                if not self.wait_for_sap(30):
                    print(f"Timeout durante l'esecuzione della transazione")
                    return False
                time.sleep(0.5)
                # Attendi che la clipboard sia riempita
                if not self.wait_for_clipboard_data(30):
                    # Gestisci il caso in cui non sono stati trovati dati
                    print("Nessun dato trovato nella clipboard")
                    # Eventuali azioni di fallback
                # Leggo il contenuto della clipboard
                return True, self.clipboard_data()
            
        except Exception as e:
            #print(f"Errore nell'estrazione ZPMR_CONTROL_FL1: {str(e)}")
            return False, str(e)


    def extract_ZPMR_CONTROL_FL2(self, fltechnology: str) -> List[Dict]:
        """
        Estrae dati relativi alla tabella ZPMR_CTRL_ASS utilizzando la transazione SE16
        
        Args:
            fltechnology: Tecnologia ricavate dalle FL
            
        Returns:
            True se la transazione va a buon fine, False altrimenti
        """
        try:
            # Svuota la clipboard prima dell'estrazione
            win32clipboard.OpenClipboard()
            win32clipboard.EmptyClipboard()
            win32clipboard.CloseClipboard()
            # Naviga alla transazione SE16
            self.session.findById("wnd[0]/tbar[0]/okcd").text = "/nSE16"
            self.session.findById("wnd[0]").sendVKey(0)
            self.session.findById("wnd[0]/usr/ctxtDATABROWSE-TABLENAME").text = "ZPMR_CONTROL_FL2"
            self.session.findById("wnd[0]").sendVKey(0)
            time.sleep(0.5)
            self.session.findById("wnd[0]/usr/ctxtI2-LOW").text = "Z-R" + fltechnology + "S"
            self.session.findById("wnd[0]/usr/ctxtI4-LOW").text = fltechnology
            self.session.findById("wnd[0]/usr/txtMAX_SEL").text = "9999999"
            self.session.findById("wnd[0]/usr/ctxtI4-LOW").setFocus()
            self.session.findById("wnd[0]/usr/ctxtI4-LOW").caretPosition = 1
            self.session.findById("wnd[0]/tbar[1]/btn[8]").press()
            # Attendi che SAP sia pronto
            if not self.wait_for_sap(30):
                print(f"Timeout durante l'esecuzione della transazione")
                return False
            time.sleep(0.5)
            self.session.findById("wnd[0]/mbar/menu[0]/menu[10]/menu[3]/menu[2]").select()
            self.session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[4,0]").select()
            self.session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[4,0]").setFocus()
            self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
            # Attendi che SAP sia pronto
            if not self.wait_for_sap(30):
                print(f"Timeout durante l'esecuzione della transazione")
                return False
            time.sleep(0.5)
            # Attendi che la clipboard sia riempita
            if not self.wait_for_clipboard_data(30):
                # Gestisci il caso in cui non sono stati trovati dati
                print("Nessun dato trovato nella clipboard")
                # Eventuali azioni di fallback
            # Leggo il contenuto della clipboard
            return self.clipboard_data()
            
        except Exception as e:
            print(f"Errore nell'estrazione ZPMR_CONTROL_FL2: {str(e)}")
            return False
        

    def extract_ZPMR_CTRL_ASS(self, fltechnology: str) -> List[Dict]:
        """
        Estrae dati relativi alla tabella ZPMR_CTRL_ASS utilizzando la transazione SE16
        
        Args:
            fltechnology: Tecnologia ricavate dalle FL
            
        Returns:
            True se la transazione va a buon fine, False altrimenti
        """
        try:
            # Svuota la clipboard prima dell'estrazione
            win32clipboard.OpenClipboard()
            win32clipboard.EmptyClipboard()
            win32clipboard.CloseClipboard()
            # Naviga alla transazione SE16
            self.session.findById("wnd[0]/tbar[0]/okcd").text = "/nSE16"
            self.session.findById("wnd[0]").sendVKey(0)
            self.session.findById("wnd[0]/usr/ctxtDATABROWSE-TABLENAME").text = "ZPMR_CTRL_ASS"
            self.session.findById("wnd[0]").sendVKey(0)
            time.sleep(0.5)
            # filtro in base alla tecnologia                
            self.session.findById("wnd[0]/usr/txtI4-LOW").text = "Z-R" + fltechnology + "S"
            self.session.findById("wnd[0]/usr/txtI5-LOW").text = fltechnology      
            # modifico il numero massimo di risultati
            self.session.findById("wnd[0]/usr/txtMAX_SEL").text = "9999999"
            self.session.findById("wnd[0]").sendVKey(0)
            # avvio la transazione
            self.session.findById("wnd[0]").sendVKey(8)
            # Attendi che SAP sia pronto
            if not self.wait_for_sap(30):
                print(f"Timeout durante l'esecuzione della transazione")
                return False
            time.sleep(0.5)    
            # esporto i valori nella clipboard
            self.session.findById("wnd[0]/mbar/menu[0]/menu[10]/menu[3]/menu[2]").select()
            # Attendi che SAP sia pronto
            if not self.wait_for_sap(30):
                print(f"Timeout durante l'esecuzione della transazione")
                return False
            time.sleep(0.5)                          
            self.session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[4,0]").select()
            self.session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[4,0]").setFocus()
            self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
            # Attendi che SAP sia pronto
            if not self.wait_for_sap(30):
                print(f"Timeout durante l'esecuzione della transazione")
                return False
            time.sleep(0.5)
            # Attendi che la clipboard sia riempita
            if not self.wait_for_clipboard_data(30):
                # Gestisci il caso in cui non sono stati trovati dati
                print("Nessun dato trovato nella clipboard")
                # Eventuali azioni di fallback
            # Leggo il contenuto della clipboard
            return self.clipboard_data()
            
        except Exception as e:
            print(f"Errore nell'estrazione ZPMR_CTRL_ASS: {str(e)}")
            return False        
# ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

    def extract_ZPM4R_GL_T_FL(self, fltechnology: str) -> List[Dict]:
        """
        Estrae dati relativi alla tabella ZPM4R_GL_T_FL utilizzando la transazione SE16
        
        Args:
            fltechnology: Tecnologia ricavate dalle FL
            
        Returns:
            True se la transazione va a buon fine, False altrimenti
        """
        try:
            # Svuota la clipboard prima dell'estrazione
            win32clipboard.OpenClipboard()
            win32clipboard.EmptyClipboard()
            win32clipboard.CloseClipboard()
            # Naviga alla transazione SE16
            self.session.findById("wnd[0]/tbar[0]/okcd").text = "/nSE16"
            self.session.findById("wnd[0]").sendVKey(0)
            self.session.findById("wnd[0]/usr/ctxtDATABROWSE-TABLENAME").text = "ZPM4R_GL_T_FL"
            self.session.findById("wnd[0]").sendVKey(0)
            time.sleep(0.5)
            # filtro in base alla tecnologia                
            self.session.findById("wnd[0]/usr/ctxtI4-LOW").text = "Z-R" + fltechnology + "S"
            self.session.findById("wnd[0]/usr/ctxtI5-LOW").text = fltechnology    
            # modifico il numero massimo di risultati
            self.session.findById("wnd[0]/usr/txtMAX_SEL").text = "9999999"
            self.session.findById("wnd[0]").sendVKey(0)
            # avvio la transazione
            self.session.findById("wnd[0]").sendVKey(8)
            # Attendi che SAP sia pronto
            if not self.wait_for_sap(30):
                print(f"Timeout durante l'esecuzione della transazione")
                return False
            time.sleep(0.5)    
            # esporto i valori nella clipboard
            self.session.findById("wnd[0]/mbar/menu[0]/menu[10]/menu[3]/menu[2]").select()
            # Attendi che SAP sia pronto
            if not self.wait_for_sap(30):
                print(f"Timeout durante l'esecuzione della transazione")
                return False
            time.sleep(0.5)                          
            self.session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[4,0]").select()
            self.session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[4,0]").setFocus()
            self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
            # Attendi che SAP sia pronto
            if not self.wait_for_sap(30):
                print(f"Timeout durante l'esecuzione della transazione")
                return False
            time.sleep(0.5)
            # Attendi che la clipboard sia riempita
            if not self.wait_for_clipboard_data(30):
                # Gestisci il caso in cui non sono stati trovati dati
                print("Nessun dato trovato nella clipboard")
                # Eventuali azioni di fallback
            # Leggo il contenuto della clipboard
            return self.clipboard_data()
            
        except Exception as e:
            print(f"Errore nell'estrazione ZPM4R_GL_T_FL: {str(e)}")
            return False
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
            
    def wait_for_clipboard_data(self, timeout: int = 30) -> bool:
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
                win32clipboard.OpenClipboard()
                try:
                    # Verifica se c'è del testo nella clipboard
                    if win32clipboard.IsClipboardFormatAvailable(win32clipboard.CF_UNICODETEXT):
                        data = win32clipboard.GetClipboardData(win32clipboard.CF_UNICODETEXT)
                        if data and data.strip():
                            print("Dati trovati nella clipboard")
                            return True
                finally:
                    win32clipboard.CloseClipboard()
                
                # Stampa il messaggio di attesa solo ogni print_interval secondi
                if current_time - last_print_time >= print_interval:
                    print("In attesa dei dati nella clipboard...")
                    last_print_time = current_time
                
                # Aspetta prima del prossimo controllo
                time.sleep(0.1)  # Ridotto il tempo di attesa per una risposta più veloce
                
            except win32clipboard.error as we:
                print(f"Errore Windows Clipboard: {str(we)}")
                time.sleep(0.5)  # Attesa più lunga in caso di errore
                continue
            except Exception as e:
                print(f"Errore durante il controllo della clipboard: {str(e)}")
                return False  

    def clipboard_data(self) -> Optional[pd.DataFrame]:
        """
        Legge i dati dalla clipboard, rimuove le righe di separazione e le colonne vuote,
        e gestisce le intestazioni duplicate.
        
        Returns:
            DataFrame Pandas pulito o None in caso di errore
        """
        try:
            # Legge il contenuto della clipboard
            win32clipboard.OpenClipboard()
            try:
                data = win32clipboard.GetClipboardData(win32clipboard.CF_UNICODETEXT)
            finally:
                win32clipboard.CloseClipboard()

            if not data:
                print("Nessun dato trovato nella clipboard")
                return None
            else:
                 return data

        except Exception as e:
            print(f"Errore durante lettura dei dati dalla clipboard: {str(e)}")
            return None

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