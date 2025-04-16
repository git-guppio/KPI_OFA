import os

# Ottieni il percorso assoluto della directory contenente lo script principale
A_ScriptDir = os.path.dirname(os.path.dirname(os.path.abspath(__file__))) # salgo di un livello rispetto alla cartella dove è contenuto il file constant.py

print(f"La directory dello script è: {A_ScriptDir}")

# ----------------------------------------------------
# Debug mode
# ----------------------------------------------------
DEBUG_MODE = False

# ----------------------------------------------------
# Intestazioni per i file di upload
# ----------------------------------------------------
# Intestazione dei file per upload tabelle di controllo globali
intestazione_ZPMR_FL_2 = "TPLKZ;FLTYP;FLLEVEL;LAND1;VALUE"
intestazione_ZPMR_FL_n = "TPLKZ;FLTYP;FLLEVEL;VALUE;VALUETX;REFLEVEL"
# Intestazione dei file per upload CTRL_ASS e TECH_OBJ
intestazione_CTRL_ASS = "VALUE;SUB_VALUE;SUB_VALUE2;TPLKZ;FLTYP;FLLEVEL;CODE_SEZ_PM;CODE_SIST;CODE_PARTE;TIPO_ELEM"
intestazione_TECH_OBJ = "VALUE;SUB_VALUE;SUB_VALUE2;TPLKZ;FLTYP;FLLEVEL;EQART;RBNR;NATURE;DOCUMENTO_PM"

# ----------------------------------------------------
# Definizione dei percorsi dei file per il controllo delle FL
# ----------------------------------------------------
file_Country = os.path.join(A_ScriptDir, "Config", "country.csv")
file_Tech = os.path.join(A_ScriptDir, "Config", "Technology.csv")
file_Rules = os.path.join(A_ScriptDir, "Config", "Rules.csv")
file_Mask = os.path.join(A_ScriptDir, "Config", "Mask_FL.csv")

# ----------------------------------------------------
# Definizione dei percorsi dei file Guideline per il controllo delle FL
# ----------------------------------------------------
file_FL_Wind = os.path.join(A_ScriptDir, "Config", "Wind_FL_GuideLine.csv")
file_FL_Bess = os.path.join(A_ScriptDir, "Config", "Bess_FL_GuideLine.csv")
file_FL_Solar_Common = os.path.join(A_ScriptDir, "Config", "Solar_FL_Common_GuideLine.csv")
file_FL_W_SubStation = os.path.join(A_ScriptDir, "Config", "Wind_FL_SubStation_Guideline.csv")
file_FL_S_SubStation = os.path.join(A_ScriptDir, "Config", "Solar_FL_SubStation_Guideline.csv")
file_FL_B_SubStation = os.path.join(A_ScriptDir, "Config", "Bess_FL_SubStation_Guideline.csv")
file_FL_Solar_CentralInv = os.path.join(A_ScriptDir, "Config", "Solar_FL_CentrealInv_GuideLine.csv")
file_FL_Solar_StringInv = os.path.join(A_ScriptDir, "Config", "Solar_FL_StringInv_GuideLine.csv")
file_FL_Solar_InvModule = os.path.join(A_ScriptDir, "Config", "Solar_FL_InvModule_GuideLine.csv")
file_FL_Hydro = os.path.join(A_ScriptDir, "Config", "Hydro_FL_GuideLine.csv")

# ----------------------------------------------------
# Definizione del percorsi dei file per l'upload
# ----------------------------------------------------
path_file_UpLoad = os.path.join(A_ScriptDir, "FileUpLoad")
# Definizione dei files per l'upload delle tabelle globali
file_ZPMR_FL_2_UpLoad = os.path.join(A_ScriptDir, "FileUpLoad", "ZPMR_FL_2_UpLoad.csv")
file_ZPMR_FL_n_UpLoad = os.path.join(A_ScriptDir, "FileUpLoad", "ZPMR_FL_n_UpLoad.csv")
# Definizione del file per l'upload delle tabelle Control Asset
file_ZPMR_CTRL_ASS_UpLoad = os.path.join(A_ScriptDir, "FileUpLoad", "ZPMR_CTRL_ASS_UpLoad.csv")
# Definizione del file per l'upload delle tabelle Technical Object
file_ZPMR_TECH_OBJ_UpLoad = os.path.join(A_ScriptDir, "FileUpLoad", "ZPMR_TECH_OBJ_UpLoad.csv")
# Definizione del file per il controllo delle Technical Object di tutte le FL
file_Check_TECH_OBJ = os.path.join(A_ScriptDir, "FileUpLoad", "FL_Check_TECH_OBJ.csv")

# timeout operazioni in SAP
timeoutSeconds = 30
# Nomi colonne tabella CTRL_ASS - Control Asset
CTRL_ASS_Valore_Livello = "Valore Livello"
CTRL_ASS_Valore_Liv_Superiore_1 = "Valore Liv. Superiore"
CTRL_ASS_Valore_Liv_Superiore_2 = "Valore Liv. Superiore"
CTRL_ASS_LivelloSedeTecnica = "Liv.Sede"