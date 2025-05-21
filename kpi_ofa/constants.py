import os

# Ottieni il percorso assoluto della directory contenente lo script principale
A_ScriptDir = os.path.dirname(os.path.dirname(os.path.abspath(__file__))) # salgo di un livello rispetto alla cartella dove è contenuto il file constant.py

print(f"La directory dello script è: {A_ScriptDir}")

# ----------------------------------------------------
# File salvataggio configurazione json
# ----------------------------------------------------
configuration_json = os.path.join(A_ScriptDir, "config.json")

# ----------------------------------------------------
# Debug mode
# ----------------------------------------------------
DEBUG_MODE = False

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

default_config = {
    "save_directory": os.path.expanduser("~"),
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

# timeout operazioni in SAP
timeoutSeconds = 30
