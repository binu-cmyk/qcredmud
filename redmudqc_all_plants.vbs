' ============================================================
' RedMud QC - download ZQM_01 for every plant in "plant code",
' then merge all plant files into one Excel workbook.
'
' How to run: log in to SAP GUI, then double-click this file.
' ============================================================
Option Explicit

' ---- Settings ----
Const TCODE        = "zqm_01"
Const MATERIAL     = "*redmud*"
Const DATE_FROM    = "01.04.2025"
Const DATE_TO      = "30.09.2026"
Const OUT_DIR      = "D:\OneDrive - Aditya Birla Group\CPC- Raw Materials - Documents\Binu\Binu RM\Redmud\redmudqcdump\"
Const MERGED_FILE  = "RedMud_QC_All_Plants.xlsx"
Const RUN_DOWNLOAD = True   ' False = only merge the files already downloaded
Const RUN_MERGE    = True

Dim fso, scriptDir, plants, p, session, results
Set fso = CreateObject("Scripting.FileSystemObject")
scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)

plants = LoadPlants(scriptDir & "\plant code")
If UBound(plants) < 0 Then
    MsgBox "No plant codes found in:" & vbCrLf & scriptDir & "\plant code", vbExclamation, "RedMud QC"
    WScript.Quit
End If
If Not fso.FolderExists(OUT_DIR) Then
    MsgBox "Output folder does not exist:" & vbCrLf & OUT_DIR, vbExclamation, "RedMud QC"
    WScript.Quit
End If

results = ""

If RUN_DOWNLOAD Then
    Set session = ConnectSap()
    If session Is Nothing Then
        MsgBox "Could not connect to SAP GUI. Open SAP, log in, and enable scripting.", vbExclamation, "RedMud QC"
        WScript.Quit
    End If
    session.findById("wnd[0]").maximize

    For Each p In plants
        results = results & p(0) & " (" & p(1) & "): " & ExportPlant(p(0)) & vbCrLf
    Next

    ' Back to SAP main menu
    On Error Resume Next
    session.findById("wnd[0]/tbar[0]/okcd").text = "/n"
    session.findById("wnd[0]").sendVKey 0
    On Error GoTo 0

    CloseSapExcelWindows
End If

If RUN_MERGE Then results = results & vbCrLf & MergeFiles()

MsgBox results, vbInformation, "RedMud QC - finished"


' ------------------------------------------------------------
' Reads "plant code" (Name<TAB>Code per line, blank lines ignored)
' Returns an array of Array(code, name)
' ------------------------------------------------------------
Function LoadPlants(path)
    Dim list, n, ts, line, parts, code, name
    list = Array()
    n = 0
    If fso.FileExists(path) Then
        Set ts = fso.OpenTextFile(path, 1)
        Do Until ts.AtEndOfStream
            line = Trim(Replace(ts.ReadLine, vbCr, ""))
            parts = Split(line, vbTab)
            If UBound(parts) = 0 Then parts = Split(line, " ")
            If UBound(parts) >= 0 Then
                code = UCase(Trim(parts(UBound(parts))))
                name = Trim(parts(0))
                If code <> "" Then
                    ReDim Preserve list(n)
                    list(n) = Array(code, name)
                    n = n + 1
                End If
            End If
        Loop
        ts.Close
    End If
    LoadPlants = list
End Function

Function ConnectSap()
    Dim sapGui, sapApp, conn
    Set ConnectSap = Nothing
    On Error Resume Next
    Set sapGui = GetObject("SAPGUI")
    Set sapApp = sapGui.GetScriptingEngine
    Set conn = sapApp.Children(0)
    Set ConnectSap = conn.Children(0)
    If Err.Number <> 0 Then Set ConnectSap = Nothing
    On Error GoTo 0
End Function

' ------------------------------------------------------------
' Runs ZQM_01 for one plant and exports it to OUT_DIR\<plant>.xlsx
' Returns "OK", "NO DATA - ..." or "FAILED - ..."
' ------------------------------------------------------------
Function ExportPlant(code)
    Dim fileName, filePath, t
    fileName = LCase(code) & ".xlsx"
    filePath = OUT_DIR & fileName

    ' Remove the previous download so old data is never merged
    CloseWorkbookByName fileName
    On Error Resume Next
    If fso.FileExists(filePath) Then fso.DeleteFile filePath, True
    If Err.Number <> 0 Then
        ExportPlant = "FAILED - old file is open/locked: " & filePath
        Exit Function
    End If

    ' Selection screen (fresh start each plant, so one failure does not break the rest)
    ClosePopups
    session.findById("wnd[0]/tbar[0]/okcd").text = "/n" & TCODE
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/usr/ctxtP_WERKS").text = code
    session.findById("wnd[0]/usr/ctxtS_MATNR-LOW").text = MATERIAL
    session.findById("wnd[0]/usr/ctxtS_DATUM-LOW").text = DATE_FROM
    session.findById("wnd[0]/usr/ctxtS_DATUM-HIGH").text = DATE_TO
    session.findById("wnd[0]/tbar[1]/btn[8]").press
    If Err.Number <> 0 Then
        ExportPlant = "FAILED - " & Err.Description
        ClosePopups
        Exit Function
    End If

    ' No data: SAP shows a message popup and/or stays on the selection screen
    If Not session.findById("wnd[1]", False) Is Nothing Then session.findById("wnd[1]").sendVKey 0
    If Not session.findById("wnd[0]/usr/ctxtP_WERKS", False) Is Nothing Then
        ExportPlant = "NO DATA - " & session.findById("wnd[0]/sbar").Text
        Exit Function
    End If

    ' Export to spreadsheet (same steps as the recorded script)
    session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[2]").select
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[2,0]").select
    session.findById("wnd[1]/tbar[0]/btn[0]").press
    session.findById("wnd[1]/tbar[0]/btn[20]").press
    session.findById("wnd[1]/usr/ctxtDY_PATH").text = OUT_DIR
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = fileName
    session.findById("wnd[1]/tbar[0]/btn[0]").press
    If Err.Number <> 0 Then
        ExportPlant = "FAILED during export - " & Err.Description
        ClosePopups
        Exit Function
    End If
    On Error GoTo 0

    ' Wait up to 30 s for the file to appear
    t = 0
    Do While Not fso.FileExists(filePath) And t < 60
        WScript.Sleep 500
        t = t + 1
    Loop
    If fso.FileExists(filePath) Then
        ExportPlant = "OK"
    Else
        ExportPlant = "FAILED - file was not created"
    End If
End Function

Sub ClosePopups()
    Dim i
    On Error Resume Next
    For i = 1 To 5
        If session.findById("wnd[1]", False) Is Nothing Then Exit For
        session.findById("wnd[1]").sendVKey 12   ' F12 = Cancel
    Next
End Sub

' SAP opens each exported file in Excel; close those so they are not locked
Sub CloseWorkbookByName(fileName)
    Dim xl, i
    On Error Resume Next
    Set xl = GetObject(, "Excel.Application")
    If Err.Number <> 0 Then Exit Sub
    For i = xl.Workbooks.Count To 1 Step -1
        If LCase(xl.Workbooks(i).Name) = LCase(fileName) Then xl.Workbooks(i).Close False
    Next
End Sub

Sub CloseSapExcelWindows()
    Dim xl, pl
    WScript.Sleep 5000   ' give Excel time to open the last file
    For Each pl In plants
        CloseWorkbookByName LCase(pl(0)) & ".xlsx"
    Next
    On Error Resume Next
    Set xl = GetObject(, "Excel.Application")
    If Err.Number = 0 Then
        If xl.Workbooks.Count = 0 Then xl.Quit
    End If
End Sub

' ------------------------------------------------------------
' Merges OUT_DIR\<plant>.xlsx files into OUT_DIR\MERGED_FILE
' Adds "Plant Code" and "Plant Name" columns in front.
' ------------------------------------------------------------
Function MergeFiles()
    Dim xl, outWb, outWs, wb, ws, ur, pl, f, lastRow, lastCol
    Dim headerCols, nextRow, n, plantsMerged, mergedPath, warn, msg
    mergedPath = OUT_DIR & MERGED_FILE
    CloseWorkbookByName MERGED_FILE

    Set xl = CreateObject("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    xl.ScreenUpdating = False
    Set outWb = xl.Workbooks.Add(-4167)   ' one empty sheet
    Set outWs = outWb.Worksheets(1)
    outWs.Name = "All Plants"

    headerCols = 0
    nextRow = 2
    plantsMerged = 0
    warn = ""

    For Each pl In plants
        f = OUT_DIR & LCase(pl(0)) & ".xlsx"
        If fso.FileExists(f) Then
            Set wb = xl.Workbooks.Open(f, 0, True)
            Set ws = wb.Worksheets(1)
            Set ur = ws.UsedRange
            lastRow = ur.Row + ur.Rows.Count - 1
            lastCol = ur.Column + ur.Columns.Count - 1

            If headerCols = 0 Then
                outWs.Cells(1, 1).Value = "Plant Code"
                outWs.Cells(1, 2).Value = "Plant Name"
                ws.Range(ws.Cells(1, 1), ws.Cells(1, lastCol)).Copy outWs.Cells(1, 3)
                headerCols = lastCol
            ElseIf lastCol <> headerCols Then
                warn = warn & "  " & pl(0) & " has " & lastCol & " columns (expected " & headerCols & ")" & vbCrLf
            End If

            If lastRow >= 2 Then
                n = lastRow - 1
                ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, lastCol)).Copy outWs.Cells(nextRow, 3)
                outWs.Range(outWs.Cells(nextRow, 1), outWs.Cells(nextRow + n - 1, 1)).Value = pl(0)
                outWs.Range(outWs.Cells(nextRow, 2), outWs.Cells(nextRow + n - 1, 2)).Value = pl(1)
                nextRow = nextRow + n
                plantsMerged = plantsMerged + 1
            End If
            wb.Close False
        End If
    Next

    If headerCols = 0 Then
        outWb.Close False
        xl.Quit
        MergeFiles = "MERGE: no plant files found in " & OUT_DIR
        Exit Function
    End If

    On Error Resume Next
    outWs.Range(outWs.Cells(1, 1), outWs.Cells(1, headerCols + 2)).Font.Bold = True
    outWs.UsedRange.AutoFilter
    outWs.UsedRange.Columns.AutoFit
    Err.Clear

    If fso.FileExists(mergedPath) Then fso.DeleteFile mergedPath, True
    outWb.SaveAs mergedPath, 51   ' 51 = .xlsx
    If Err.Number <> 0 Then
        msg = "MERGE FAILED - could not save " & mergedPath & " (" & Err.Description & ")"
    Else
        msg = "MERGED " & plantsMerged & " plants, " & (nextRow - 2) & " rows ->" & vbCrLf & mergedPath
    End If
    If warn <> "" Then msg = msg & vbCrLf & vbCrLf & "Warning - column count differs:" & vbCrLf & warn
    On Error GoTo 0

    outWb.Close False
    xl.Quit
    MergeFiles = msg
End Function
