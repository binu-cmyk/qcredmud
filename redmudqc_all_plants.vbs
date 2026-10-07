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

path      = "D:\OneDrive - Aditya Birla Group\CPC- Raw Materials - Documents\Binu\Binu RM\Redmud\redmudqcdump\"
plantFile = "D:\RedMud QC\plant code"
mergeFile = path & "RedMud_QC_All_Plants.xlsx"

Set fso = CreateObject("Scripting.FileSystemObject")

'---- read plant codes (2nd column of plant code file) ----
plants = ""
Set ts = fso.OpenTextFile(plantFile, 1)
Do Until ts.AtEndOfStream
   line = Split(ts.ReadLine, vbTab)
   If UBound(line) >= 1 Then
      If Trim(line(1)) <> "" Then plants = plants & LCase(Trim(line(1))) & ","
   End If
Loop
ts.Close
plants = Split(plants, ",")

'---- download plantwise ----
session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "/nzqm_01"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtS_MATNR-LOW").text = "*redmud*"
session.findById("wnd[0]/usr/ctxtS_DATUM-LOW").text = "01.04.2025"
session.findById("wnd[0]/usr/ctxtS_DATUM-HIGH").text = "30.09.2026"

For i = 0 To UBound(plants) - 1
   plant = plants(i)
   If fso.FileExists(path & plant & ".xlsx") Then fso.DeleteFile path & plant & ".xlsx", True

   session.findById("wnd[0]/usr/ctxtP_WERKS").text = plant
   session.findById("wnd[0]/usr/ctxtP_WERKS").caretPosition = 4
   session.findById("wnd[0]/tbar[1]/btn[8]").press

   'no data -> close message popup, stay on selection screen, go to next plant
   If Not session.findById("wnd[1]", False) Is Nothing Then session.findById("wnd[1]").sendVKey 0

   If session.findById("wnd[0]/usr/ctxtP_WERKS", False) Is Nothing Then
      session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[2]").select
      session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[2,0]").select
      session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[2,0]").setFocus
      session.findById("wnd[1]/tbar[0]/btn[0]").press
      session.findById("wnd[1]/tbar[0]/btn[20]").press
      session.findById("wnd[1]/usr/ctxtDY_PATH").text = path
      session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = plant & ".xlsx"
      session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 9
      session.findById("wnd[1]/tbar[0]/btn[0]").press
      session.findById("wnd[0]/tbar[0]/btn[3]").press
   End If
Next

'---- merge all plant files into one ----
Set xl = CreateObject("Excel.Application")
xl.DisplayAlerts = False
Set outWb = xl.Workbooks.Add(-4167)
Set outWs = outWb.Sheets(1)
outWs.Name = "All Plants"
outRow = 1

For i = 0 To UBound(plants) - 1
   plant = plants(i)
   If fso.FileExists(path & plant & ".xlsx") Then
      Set wb = xl.Workbooks.Open(path & plant & ".xlsx", 0, True)
      Set ws = wb.Sheets(1)
      lastRow = ws.UsedRange.Rows.Count
      lastCol = ws.UsedRange.Columns.Count

      'header only once, from first file
      If outRow = 1 Then
         outWs.Cells(1, 1).Value = "Plant"
         ws.Range(ws.Cells(1, 1), ws.Cells(1, lastCol)).Copy outWs.Cells(1, 2)
         outRow = 2
      End If

      If lastRow > 1 Then
         ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, lastCol)).Copy outWs.Cells(outRow, 2)
         outWs.Range(outWs.Cells(outRow, 1), outWs.Cells(outRow + lastRow - 2, 1)).Value = UCase(plant)
         outRow = outRow + lastRow - 1
      End If
      wb.Close False
   End If
Next

If fso.FileExists(mergeFile) Then fso.DeleteFile mergeFile, True
outWb.SaveAs mergeFile, 51
outWb.Close False
xl.Quit

MsgBox "Done. Merged file:" & vbCrLf & mergeFile
