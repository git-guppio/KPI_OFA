import win32com.client

def map_field_names(id_grid_control, session):
    """
    Crea un dizionario che mappa i field name delle colonne 
    a tutte le loro possibili varianti testuali.
    
    Args:
        id_grid_control: ID del controllo griglia in SAP (es. "wnd[0]/usr/cntlGRID1/shellcont/shell")
        session: oggetto sessione SAP attiva
        
    Returns:
        dict: dizionario con field_name come chiave e lista di varianti come valore
    """
    field_names_map = {}
    
    try:
        # Accedi al GridView
        grid_view = session.findById(id_grid_control)
        
        # Ricava tutti i field name delle colonne presenti
        columns_field_names = grid_view.ColumnOrder
        
        # Per ogni field name, ricava tutti i titoli possibili
        for index, field_name in enumerate(columns_field_names):
            # Ottieni il titolo visualizzato correntemente
            column_title = grid_view.GetDisplayedColumnTitle(field_name)
            
            # Ottieni tutte le varianti testuali per questo campo
            # IMPORTANTE: converti subito in lista
            all_titles = list(grid_view.GetColumnTitles(field_name))
            
            # Rimuovi eventuali duplicati mantenendo l'ordine
            unique_titles = []
            for title in all_titles:
                if title and title not in unique_titles:  # Ignora valori vuoti e duplicati
                    unique_titles.append(title)
            
            # Memorizza nel dizionario
            field_names_map[field_name] = unique_titles
            
            print(f"{field_name} - Titolo corrente: {column_title}")
            print(f"  Tutte le varianti: {unique_titles}")
        
        return field_names_map
        
    except Exception as e:
        print(f"Errore durante l'estrazione dei field names: {e}")
        return None


def get_sap_session():
    """
    Si connette a una sessione SAP esistente.
    
    Returns:
        oggetto sessione SAP
    """
    try:
        # Connessione al SAP GUI Scripting
        sap_gui_auto = win32com.client.GetObject("SAPGUI")
        application = sap_gui_auto.GetScriptingEngine
        
        # Ottieni la connessione attiva (prima disponibile)
        connection = application.Children(0)
        
        # Ottieni la sessione attiva (prima disponibile)
        session = connection.Children(0)
        
        print(f"Connesso a SAP - Sessione: {session.Info.Transaction}")
        return session
        
    except Exception as e:
        print(f"Errore nella connessione a SAP: {e}")
        return None


# ===== UTILIZZO =====
if __name__ == "__main__":
    # Connetti alla sessione SAP
    session = get_sap_session()
    
    if session:
        # ID del controllo griglia
        grid_id = "wnd[0]/usr/cntlGRID1/shellcont/shell"
        
        # Estrai la mappatura
        field_map = map_field_names(grid_id, session)
        
        if field_map:
            print("\n=== DIZIONARIO COMPLETO ===")
            for field, variants in field_map.items():
                print(f"\n{field}:")
                for variant in variants:
                    print(f"  - {variant}")