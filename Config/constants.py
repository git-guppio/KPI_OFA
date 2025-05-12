import os

# Ottieni il percorso assoluto della directory contenente lo script principale
A_ScriptDir = os.path.dirname(os.path.dirname(os.path.abspath(__file__))) # salgo di un livello rispetto alla cartella dove è contenuto il file constant.py

print(f"La directory dello script è: {A_ScriptDir}")

# ----------------------------------------------------
# Debug mode
# ----------------------------------------------------
DEBUG_MODE = False

# ----------------------------------------------------
# Dati relativi al file Excel con i dettagli OFA
# ----------------------------------------------------
# Nome dello sheet del file excel
required_sheet = "OnFieldApp.trace_kpi_user_actio"
# Definisci le colonne richieste
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
    }
}

# timeout operazioni in SAP
timeoutSeconds = 30
