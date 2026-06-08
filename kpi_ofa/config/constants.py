import os

# Ottieni il percorso assoluto della directory contenente lo script principale
A_ScriptDir = os.path.dirname(os.path.dirname(os.path.abspath(__file__))) # salgo di un livello rispetto alla cartella dove è contenuto il file constant.py
print(f"Le directory di lavoro è:\n - directory principale: {A_ScriptDir}")

# ----------------------------------------------------
# File salvataggio configurazione json
# ----------------------------------------------------
configuration_json = os.path.join(A_ScriptDir, "config", "config.json")

# ----------------------------------------------------
# Debug mode
# ----------------------------------------------------
#DEBUG_MODE = True
DEBUG_MODE = False
# Mapping tra file e attributi DataFrame
test_file_to_df_mapping = {
    "df_AFKO_test.xlsx": "df_AFKO",
    "df_excel_norm_test.xlsx": "df_excel_normalized", 
    "IW29_AdM_test.xlsx": "df_IW29",
    "IW39_OdM_test.xlsx": "df_IW39",
    "plants.xlsx": "df_plants"
}


# ----------------------------------------------------
# Dati relativi al file Excel con i dettagli OFA
# ----------------------------------------------------
# Nome dello sheet del file excel
required_sheet = "OnFieldApp.trace_kpi_user_actio"
# Definisci le colonne richieste nel file Excel
required_columns = [
    "user", 
    "creationDate", 
    "country", 
    "tecnology", 
    "action", 
    "idItem", 
    "functionalLocation"
]

# ----------------------------------------------------
# Coonfigurazione di default che viene caricata nel caso non esista il file json
# ----------------------------------------------------
default_config = {
    "data_directory": os.path.join(A_ScriptDir, "data"),

    "technologies": {
        "BESS": ["ITE", "USE", "CLE"],
        "SOLAR": ["ITS", "USS", "CLS", "BRS", "COS", "MXS", "PAS", "ZAS", "ESS", "ZMS"],
        "WIND": ["ITW", "USW", "CLW", "BRW", "CAW", "MXW", "ZAW", "ESW"]
    },

    "operations": {
        "estrai_AdM": True, 
        "estrai_OdM": True,
        "estrai_AFKO": True,        
        "elabora_xls": True
    }   
}

SAP_tech_field_names_map = {
    # Nuova versione di SAP, con nuove intestazioni
    "IW29": {
        "AEDAT": ["Data modifica", "Mod.", "Mod. il", "Dta mod."],
        "AUFNR": ["Ordine"],
        "COUNTRY": ["Codice paese", "Paese", "Pse", "Chiave paese"],
        "ERDAT": ["Data creazione", "Data cr.", "Il"],
        "LEGACY": ["Sistema proprietario dell'Anomalia Legac", "Sistema legacy", "Sis.Legacy", "Sistema proprietario dell'Anomalia Legacy"],
        "QMART": ["Tipo avviso", "Tp.avv.", "Tp."],
        "QMNUM": ["Avviso"],
        "QMTXT": ["Descrizione", "Descr."],
        "STTXT": ["Stato sistema", "St.sist."],
        "TPLNR": ["Sede tecnica", "Sede tecn."]
    },

    # Precedente Versione di SAP
    # "IW29": {
    #     "QMNUM": ["Avviso"],
    #     "AEDAT": ["Data modifica","Mod.","Mod. il"],
    #     "ERDAT": ["Data creazione","Data cr.","Il"],
    #     "QMTXT": ["Descrizione","Descr."],
    #     "QMART": ["Tipo avviso","Tp.avv.","Tp."],
    #     "TPLNR": ["Sede tecnica","Sede tecn."],
    #     "STTXT": ["Stato sistema","St.sist."],
    #     "AUFNR": ["Ordine"],
    #     "COUNTRY": ["Codice paese","Paese","Pse"],
    #     "LEGACY": ["Sistema proprietario dell'Anomalia Legac","Sistema legacy","Sis.Legacy","Sistema proprietario dell'Anomalia Legacy"]
    # },

    "IW39": {
        "AEDAT": ["Data modifica anagr. ordine", "Data modifica", "Data mod."],
        "AUART": ["Tipo di ordine", "Tipo ord.", "Tp."],
        "AUFNR": ["Ordine"],
        "COUNTRY": ["Codice paese", "Paese", "Pse", "Chiave paese"],
        "ERDAT": ["Data di acquisizione", "Data acquis.", "Data acq."],
        "GLTRP": ["Data fine cardine", "Data fine c.", "Fine card.", "Data di fine di base", "Data fine base", "Fine base"],
        "GSTRP": ["Data inizio cardine", "Data in. card.", "In. card."],
        "KTEXT": ["Tsto br.", "Descrizione", "Descr."],
        "QMNUM": ["Avviso"],
        "STTXT": ["Stato sistema", "St.sist."],
        "TPLNR": ["Sede tecnica", "Sede tecn."],
        "USTXT": ["Stato utente", "St.utente"],
        "Z_LEGACY": ["Sistema Legacy", "Sis Legacy"]
    },

    # Precedente Versione di SAP
    # "IW39": {
    #     "AUFNR": ["Ordine"],
    #     "KTEXT": ["Tsto br."],
    #     "ERDAT": ["Data di acquisizione", "Data acquis.", "Data acq."],
    #     "GSTRP": ["Data inizio cardine", "Data in. card.", "In. card."],
    #     "GLTRP": ["Data fine cardine", "Data fine c.", "Fine card."],
    #     "AEDAT": ["Data modifica anagr. ordine", "Data modifica", "Data mod."],
    #     "TPLNR": ["Sede tecnica", "Sede tecn."],
    #     "COUNTRY": ["Codice paese", "Paese", "Pse"],
    #     "AUART": ["Tipo di ordine", "Tipo ord.", "Tp."],
    #     "STTXT": ["Stato sistema", "St.sist."],
    #     "QMNUM": ["Avviso"],
    #     "Z_LEGACY": ["Sistema Legacy", "Sis Legacy"],
    #     "USTXT": ["Stato utente", "St.utente"]
    # },

    "AFKO": {
        "AUFNR": ["Ordine"],
        "GSTRP": ["Data inizio cardine", "Data di inizio base"]
    }

    # Precedente Versione di SAP
    # "AFKO": {
    #     "AUFNR": ["Ordine"],
    #     "GSTRP": ["Data inizio cardine"]
    # }
}

SAP_user_field_names_map = {
    "IW29": {
        "QMNUM": "Avviso",
        "AEDAT": "Mod. il",
        "ERDAT": "Data cr.",
        "QMTXT": "Descrizione",
        "QMART": "Tp.",
        "TPLNR": "Sede tecnica",
        "STTXT": "St.sist.",
        "AUFNR": "Ordine",
        "COUNTRY": "Pse",
        "LEGACY": "Sis.Legacy"
    },

    "IW39": {
        "AUFNR": "Ordine",
        "KTEXT": "Tsto br.",
        "ERDAT": "Data acq.",
        "GSTRP": "In. card.",
        "GLTRP": "Fine card.",
        "AEDAT": "Data mod.",
        "TPLNR": "Sede tecnica",
        "COUNTRY": "Pse",
        "AUART": "Tp.",
        "STTXT": "Stato sistema",
        "QMNUM": "Avviso",
        "Z_LEGACY": "Sis Legacy",
        "USTXT": "Stato utente"
    },

    "AFKO": {
        "AUFNR": "Ordine",
        "GSTRP": "Data inizio cardine"
    }
}

# timeout operazioni in SAP
timeoutSeconds = 30
