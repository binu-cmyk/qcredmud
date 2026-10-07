If Not IsObject(application) Then
   Set SapGuiAuto  = GetObject("SAPGUI")
   Set application = SapGuiAuto.GetScriptingEngine
End If
If Not IsObject(connection) Then
   Set connection = application.Children(0)
End If
If Not IsObject(session) Then
   Set session    = connection.Children(0)
End If
If IsObject(WScript) Then
   WScript.ConnectObject session,     "on"
   WScript.ConnectObject application, "on"
End If
session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "zqm_01"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtP_WERKS").text = "ac01"
session.findById("wnd[0]/usr/ctxtS_MATNR-LOW").text = "*redmud*"
session.findById("wnd[0]/usr/ctxtS_DATUM-LOW").text = "01.04.2025"
session.findById("wnd[0]/usr/ctxtS_DATUM-HIGH").text = "30.09.2026"
session.findById("wnd[0]/usr/ctxtP_WERKS").caretPosition = 4
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[2]").select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[2,0]").select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[2,0]").setFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ssubSUB_CONFIGURATION:SAPLSALV_GUI_CUL_EXPORT_AS:0512/cmbGS_EXPORT-DESTINATION").setFocus
session.findById("wnd[1]/tbar[0]/btn[20]").press
session.findById("wnd[1]/usr/ctxtDY_PATH").text = "D:\OneDrive - Aditya Birla Group\CPC- Raw Materials - Documents\Binu\Binu RM\Redmud\redmudqcdump\"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "ac01.xlsx"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 9
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/usr/ctxtP_WERKS").text = "ap01"
session.findById("wnd[0]/usr/ctxtP_WERKS").caretPosition = 4
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[2]").select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[2,0]").select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[2,0]").setFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/tbar[0]/btn[20]").press
session.findById("wnd[1]/usr/ctxtDY_PATH").text = "D:\OneDrive - Aditya Birla Group\CPC- Raw Materials - Documents\Binu\Binu RM\Redmud\redmudqcdump\"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "ap01.xlsx"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 9
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/usr/ctxtP_WERKS").text = "aw01"
session.findById("wnd[0]/usr/ctxtP_WERKS").caretPosition = 4
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[2]").select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[2,0]").select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[2,0]").setFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/tbar[0]/btn[20]").press
session.findById("wnd[1]/usr/ctxtDY_PATH").text = "D:\OneDrive - Aditya Birla Group\CPC- Raw Materials - Documents\Binu\Binu RM\Redmud\redmudqcdump\"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "aw01.xlsx"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 9
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press
