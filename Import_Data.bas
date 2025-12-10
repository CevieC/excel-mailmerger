Option Explicit

Sub Import_To_MailMerge()
    Dim wbDest As Workbook          ' This workbook (MailMerge / DATA workbook)
    Dim wbSrc As Workbook           ' Source workbook (ROP / COMBINED_ROP)
    Dim wsSrc As Worksheet          ' Source sheet in Workbook A
    Dim wsDest As Worksheet         ' Destination sheet in Workbook B
    
    Dim srcFilePath As Variant
    Dim lastSrcRow As Long
    Dim rowsToCopy As Long
    
    Dim srcHeaderRow As Long
    Dim srcDataStartRow As Long
    Dim destHeaderRow As Long
    Dim destDataStartRow As Long
    
    Dim oldScreenUpdating As Boolean
    Dim oldEnableEvents As Boolean
    Dim oldCalc As XlCalculation
    
    On Error GoTo ErrHandler
    
    '=============================
    ' CONFIG
    '=============================
    Const SRC_SHEET_NAME As String = "COMBINED_ROP"   ' in Workbook A
    Const DEST_SHEET_NAME As String = "DATA"          ' in THIS workbook
    
    ' Source layout (Workbook A)
    srcHeaderRow = 1          ' headers in row 1
    srcDataStartRow = 2       ' data starts row 2
    
    ' Destination layout (Workbook B)
    destHeaderRow = 1         ' headers already pre-loaded in row 1
    destDataStartRow = 2      ' paste data from row 2
    
    oldScreenUpdating = Application.ScreenUpdating
    oldEnableEvents = Application.EnableEvents
    oldCalc = Application.Calculation
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
    Set wbDest = ThisWorkbook
    
    srcFilePath = Application.GetOpenFilename( _
        FileFilter:="Excel Files (*.xlsx;*.xlsm;*.xlsb;*.xls),*.xlsx;*.xlsm;*.xlsb;*.xls", _
        Title:="Select source workbook (COMBINED_ROP)")
    
    If srcFilePath = False Then
        MsgBox "Import cancelled.", vbInformation, "Export_To_MailMerge"
        GoTo CleanExit
    End If
    
    Set wbSrc = Workbooks.Open(Filename:=CStr(srcFilePath), ReadOnly:=True)
    
    On Error Resume Next
    Set wsSrc = wbSrc.Worksheets(SRC_SHEET_NAME)
    On Error GoTo ErrHandler
    
    If wsSrc Is Nothing Then
        MsgBox "Source sheet '" & SRC_SHEET_NAME & "' not found in selected workbook.", _
               vbExclamation, "Export_To_MailMerge"
        GoTo CleanExit
    End If
    
    On Error Resume Next
    Set wsDest = wbDest.Worksheets(DEST_SHEET_NAME)
    On Error GoTo ErrHandler
    
    If wsDest Is Nothing Then
        Set wsDest = wbDest.Worksheets.Add(After:=wbDest.Worksheets(wbDest.Worksheets.Count))
        wsDest.Name = DEST_SHEET_NAME
    End If
    
    ' Clear old DATA rows but keep headers (row 1)
    wsDest.Rows(destDataStartRow & ":" & wsDest.Rows.Count).ClearContents
    
    lastSrcRow = wsSrc.Cells(wsSrc.Rows.Count, "A").End(xlUp).Row
    If lastSrcRow < srcDataStartRow Then
        MsgBox "No data found in source sheet '" & SRC_SHEET_NAME & "'.", _
               vbInformation, "Export_To_MailMerge"
        GoTo CleanExit
    End If
    
    rowsToCopy = lastSrcRow - srcDataStartRow + 1
    
    '=============================
    ' Copy A:BE (cols 1–57) → DATA A:BE
    '=============================
    wsDest.Cells(destDataStartRow, 1).Resize(rowsToCopy, 57).Value = _
        wsSrc.Cells(srcDataStartRow, 1).Resize(rowsToCopy, 57).Value
    
    '=============================
    ' Copy BQ:BR (cols 69–70) → DATA cols 58–59
    '   - so DATA structure is A:BE, then BQ:BR with no gap
    '=============================
    wsDest.Cells(destDataStartRow, 58).Resize(rowsToCopy, 2).Value = _
        wsSrc.Cells(srcDataStartRow, 69).Resize(rowsToCopy, 2).Value
    
    wsDest.Columns.AutoFit
    
    MsgBox "Data import completed successfully." & vbCrLf & _
           "Source: " & wbSrc.Name & " (" & SRC_SHEET_NAME & ")" & vbCrLf & _
           "Destination: " & wbDest.Name & " (" & DEST_SHEET_NAME & ")", _
           vbInformation, "Export_To_MailMerge"
    
CleanExit:
    On Error Resume Next
    If Not wbSrc Is Nothing Then wbSrc.Close SaveChanges:=False
    On Error GoTo 0
    
    Application.ScreenUpdating = oldScreenUpdating
    Application.EnableEvents = oldEnableEvents
    Application.Calculation = oldCalc
    Exit Sub
    
ErrHandler:
    MsgBox "Error in Export_To_MailMerge: " & Err.Description, vbExclamation, "Error"
    Resume CleanExit
End Sub
