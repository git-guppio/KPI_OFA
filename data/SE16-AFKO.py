    session.findById("wnd[0]/tbar[0]/okcd").text = "/nse16"
    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[0]/usr/ctxtDATABROWSE-TABLENAME").text = "AFKO"
    session.findById("wnd[0]").sendVKey(0)

# da incollare la lista di OdM prelevati dalla estrazione IW29
    session.findById("wnd[0]/usr/ctxtI1-LOW").text = "210000083540"
    session.findById("wnd[0]/usr/txtMAX_SEL").text = "9999999"
    session.findById("wnd[0]/usr/ctxtI1-LOW").caretPosition = 12
    session.findById("wnd[0]/tbar[1]/btn[8]").press()

# modifica Layout OFAKPIWO

    session.findById("wnd[0]/tbar[1]/btn[33]").press()
    session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").currentCellRow = 1
    session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").selectedRows = "1"
    session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").clickCurrentCell()
 # salvare in excel
