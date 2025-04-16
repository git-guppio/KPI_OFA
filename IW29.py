    session.findById("wnd[0]").resizeWorkingPane(173, 49, False)
    session.findById("wnd[0]/tbar[0]/okcd").text = "/nIW29"
    session.findById("wnd[0]").sendVKey(0)
# tutti gli stati
    session.findById("wnd[0]/usr/chkDY_OFN").selected = True
    session.findById("wnd[0]/usr/chkDY_IAR").selected = True
    session.findById("wnd[0]/usr/chkDY_RST").selected = True
    session.findById("wnd[0]/usr/chkDY_MAB").selected = True
# elimino data avviso
    session.findById("wnd[0]/usr/ctxtDATUV").text = ""
    session.findById("wnd[0]/usr/ctxtDATUB").text = ""
    session.findById("wnd[0]/usr/ctxtDATUB").setFocus()
    session.findById("wnd[0]/usr/ctxtDATUB").caretPosition = 0
    session.findById("wnd[0]/usr/btn%_QMART_%_APP_%-VALU_PUSH").press()
# tipologia avvisi
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = "Z1"
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "Z2"
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").text = "Z3"
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,3]").text = "Z4"
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,4]").text = "Z5"
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,4]").setFocus()
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,4]").caretPosition = 2
    session.findById("wnd[1]/tbar[0]/btn[8]").press()
# Sede tecnica    WSB
    session.findById("wnd[0]/usr/ctxtSTRNO-LOW").text = ""
    session.findById("wnd[0]/usr/ctxtSTRNO-LOW").setFocus()
    session.findById("wnd[0]/usr/ctxtSTRNO-LOW").caretPosition = 0
    session.findById("wnd[0]/usr/btn%_STRNO_%_APP_%-VALU_PUSH").press()
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = "++S-++++*"
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "++W-++++*"
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").text = "++E-++++*"
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,3]").text = ""
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,3]").setFocus()
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,3]").caretPosition = 0
    session.findById("wnd[1]/tbar[0]/btn[8]").press()
# Data Creazione - inizio e fine mese (da poter inserire e modificare

    session.findById("wnd[0]/usr/ctxtERDAT-LOW").text = "010325"
    session.findById("wnd[0]/usr/ctxtERDAT-HIGH").text = "310325"
    session.findById("wnd[0]/usr/ctxtERDAT-HIGH").setFocus()
    session.findById("wnd[0]/usr/ctxtERDAT-HIGH").caretPosition = 6
    session.findById("wnd[0]").sendVKey(0)

# Layout

    session.findById("wnd[0]/usr/ctxtVARIANT").text = "/KPIOFANO2"
    session.findById("wnd[0]/usr/ctxtVARIANT").setFocus()
    session.findById("wnd[0]/usr/ctxtVARIANT").caretPosition = 10
    session.findById("wnd[0]").sendVKey(0)

# Esegui

    session.findById("wnd[0]/tbar[1]/btn[8]").press()
# salvo i dati in excel
    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").contextMenu()
    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectContextMenuItem("&XXL")
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[1]/usr/ctxtDY_PATH").setFocus()
    session.findById("wnd[1]/usr/ctxtDY_PATH").caretPosition = 0
    session.findById("wnd[1]").sendVKey(4)
    session.findById("wnd[2]/usr/ctxtDY_PATH").text = "C:\\Users\\a259046\\OneDrive - Enel Spa\\SCRIPT AHK e VBA\\GITHUB\\KPI_OFA\\Dati"
    session.findById("wnd[2]/usr/ctxtDY_FILENAME").text = "export_IW29.XLSX"
    session.findById("wnd[2]/tbar[0]/btn[11]").press()
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").setFocus()
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 16
    session.findById("wnd[1]").sendVKey(4)
    session.findById("wnd[2]/tbar[0]/btn[11]").press()
    session.findById("wnd[1]/tbar[0]/btn[11]").press()
