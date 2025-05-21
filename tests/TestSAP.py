#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Script di test rapido per la GUI dell'applicazione KPI OFA.

Questo script avvia l'interfaccia grafica dell'applicazione permettendo
di testare manualmente il suo funzionamento senza dover implementare
l'integrazione con SAP o altre dipendenze esterne.
"""

import os
import sys
import re 
import time

# Aggiungi la directory principale al PYTHONPATH
current_dir = os.path.dirname(os.path.abspath(__file__))  # tests directory
project_root = os.path.dirname(current_dir)  # directory principale
sys.path.insert(0, project_root)


import json
import tempfile
import pandas as pd
import logging
from kpi_ofa.services.sap_connection import SAPGuiConnection
from kpi_ofa.services.sap_transactions import SAPDataExtractor

def SE16(session) -> tuple[bool, pd.DataFrame | None]:

    tabella_layout = session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell")
    nome_layout = "/OFAKPIWO"
    try:        
        # Ottieni il numero di righe nella tabella
        numero_righe = tabella_layout.rowCount
        
        # Cerca il layout in tutte le righe, usando la prima colonna (indice 0)
        layout_trovato = False
        for i in range(numero_righe):
            try:
                # Ottieni il valore della cella nella prima colonna
                nome_layout_riga = tabella_layout.GetCellValue(i, tabella_layout.ColumnOrder(0))
                print(f"Riga {i}: '{nome_layout_riga}'")
                
                # Se il nome corrisponde, seleziona questa riga
                if nome_layout_riga == nome_layout:
                    tabella_layout.currentCellRow = i
                    tabella_layout.selectedRows = str(i)
                    tabella_layout.clickCurrentCell()
                    print(f"Layout '{nome_layout}' trovato e selezionato")
                    layout_trovato = True
                    return True
            except Exception as e:
                print(f"Errore nella lettura della riga {i}: {str(e)}")
                continue
        
        # Gestisci il caso in cui il layout non venga trovato
        if not layout_trovato:
            print(f"Layout '{nome_layout}' non trovato nella lista dei layout disponibili")
            return False
    except Exception as e:
        print(f"Errore durante la selezione del layout: {str(e)}")
        return False

            
        

    

    except:
        print("Errore: impossibile ottenere il numero di righe e colonne")
        return False, None      

def run_test():
    """
    Avvia l'interfaccia grafica per il test manuale.
    """
    import pandas as pd

    # Lista dei valori forniti
    valori = [
        240000498399,
        210001091038,
        210001162187,
        210001162189,
        210001200428,
        210001200431,
        210001200544,
        220000073342,
        240000491458,
        240000498695,
        250000051277,
        250000051278,
        250000051279,
        250000051280,
        250000051281,
        900000047717
    ]

    # Creare un DataFrame con una colonna chiamata 'ID'
    df_OdM = pd.DataFrame(valori)

    # Visualizzare il DataFrame
    print(df_OdM)

    try:
        
        # Utilizza la factory per creare una connessione SAP
        with SAPGuiConnection() as sap:
            if sap.is_connected():
                session = sap.get_session()
                if session:
                    print("Connessione SAP attiva")
                    
                    # Crea un estrattore di dati SAP
                    extractor = SAPDataExtractor(session)
                    
                    # Estrazione dati AdM
                    print("Estrazione dati:", "info")
                    result, df = extractor.extract_SE16(df_OdM)
                    
                    if not result:
                        print("Errore: Estrazione IW29 fallita", "error")
                        return
                    
                    # Estrazione dati completata
                    print("Estrazione completata con successo", "success")
            else:
                print("Connessione SAP NON attiva", "error")
                return
            
    except Exception as e:
        print(f"Estrazione dati SAP: Errore: {str(e)}", "error")
        return

if __name__ == "__main__":
    sys.exit(run_test())