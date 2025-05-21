    session.findById("wnd[0]/tbar[0]/okcd").text = "/nIW39"
    session.findById("wnd[0]").sendVKey(0)
# tutti gli stati
# tolgo periodo
# inserisco data acquisizione

    session.findById("wnd[0]/usr/chkDY_OFN").selected = True
    session.findById("wnd[0]/usr/chkDY_IAR").selected = True
    session.findById("wnd[0]/usr/chkDY_MAB").selected = True
    session.findById("wnd[0]/usr/chkDY_HIS").selected = True
    session.findById("wnd[0]/usr/ctxtDATUV").text = ""
    session.findById("wnd[0]/usr/ctxtDATUB").text = ""
    session.findById("wnd[0]/usr/ctxtERDAT-LOW").text = "010325"
    session.findById("wnd[0]/usr/ctxtERDAT-HIGH").text = "310325"
    session.findById("wnd[0]/usr/ctxtERDAT-HIGH").setFocus()
    session.findById("wnd[0]/usr/ctxtERDAT-HIGH").caretPosition = 6
    session.findById("wnd[0]").sendVKey(0)

# sedi tecniche

    session.findById("wnd[0]/usr/btn%_STRNO_%_APP_%-VALU_PUSH").press()
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = "++S-++++*"
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "++W-++++*"
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").text = "++E-++++*"
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").setFocus()
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").caretPosition = 9
    session.findById("wnd[1]/tbar[0]/btn[8]").press()


# Layout


    session.findById("wnd[0]/usr/ctxtVARIANT").text = "/KPIOFA2"
    session.findById("wnd[0]/usr/ctxtVARIANT").setFocus()
    session.findById("wnd[0]/usr/ctxtVARIANT").caretPosition = 8
    session.findById("wnd[0]").sendVKey(0)


# avvio

    session.findById("wnd[0]/tbar[1]/btn[8]").press()

# salvo in excel
