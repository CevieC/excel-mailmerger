Option Explicit

Public Sub ImportActiveListing()
    Const DEST_SHEET As String = "AL"            
    Const DASH_SHEET As String = "Overview"      
    Const SOURCE_SHEET As String = "Setup"     
    
    Const AL_FILE_NAME_CELL As String = "S13"
    Const AL_FILE_PATH_CELL As String = "S14"
    Const AL_LAST_REFRESH_CELL As String = "S15"
    Const AL_STATUS_CELL As String = "S16"
    Const AL_ROWS_CELL As String = "S17"
    Const AL_NOTES_CELL As String = "S18"
    
    Dim wsDest As Worksheet
    Dim wsDash As Worksheet
    Dim wbSrc As Workbook
    Dim wsSrc As Worksheet
    
    Dim filePath As Variant
    Dim fileName As String
    Dim rowsLoaded As Long
    
    On Error GoTo CleanFail
    
    Set wsDest = ThisWorkbook.Sheets(DEST_SHEET)
    Set wsDash = ThisWorkbook.Sheets(DASH_SHEET)
    
    filePath = Application.GetOpenFilename( _
                    FileFilter:="Excel Files (*.xlsx;*.xlsm;*.xls),*.xlsx;*.xlsm;*.xls", _
                    Title:="Select Active Listing File")
    
    If filePath = False Then
        With wsDash
            .Range(AL_LAST_REFRESH_CELL).Value = Now
            .Range(AL_STATUS_CELL).Value = "CANCELLED"
            .Range(AL_NOTES_CELL).Value = "User cancelled Active Listing import."
        End With
        Exit Sub
    End If
    
    fileName = Dir(CStr(filePath))
    
    With wsDash
        .Range(AL_FILE_NAME_CELL).Value = fileName
        .Range(AL_FILE_PATH_CELL).Value = CStr(filePath)
        .Range(AL_LAST_REFRESH_CELL).Value = Now
        .Range(AL_STATUS_CELL).Value = "IN PROGRESS"
        .Range(AL_ROWS_CELL).Value = 0
        .Range(AL_NOTES_CELL).Value = "Importing Active Listing..."
    End With
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    Set wbSrc = Workbooks.Open(CStr(filePath))
    
    On Error Resume Next
    Set wsSrc = wbSrc.Sheets(SOURCE_SHEET)
    On Error GoTo CleanFail
    
    If wsSrc Is Nothing Then
        wbSrc.Close SaveChanges:=False
        
        With wsDash
            .Range(AL_STATUS_CELL).Value = "FAILED"
            .Range(AL_NOTES_CELL).Value = "Sheet 'Setup' not found in file."
        End With
        
        MsgBox "The selected file does not contain a sheet named 'Setup'.", _
               vbExclamation, "Active Listing Import"
        GoTo CleanExit
    End If
    
    wsDest.Cells.Clear
    wsSrc.UsedRange.Copy wsDest.Range("A1")
    
    CopyColumnWidths_AL wsSrc, wsDest
    
    rowsLoaded = wsDest.UsedRange.Rows.Count
    
    wbSrc.Close SaveChanges:=False

    With wsDash
        .Range(AL_LAST_REFRESH_CELL).Value = Now
        .Range(AL_STATUS_CELL).Value = "SUCCESS"
        .Range(AL_ROWS_CELL).Value = rowsLoaded
        .Range(AL_NOTES_CELL).Value = "OK"
    End With

CleanExit:
    Application.CutCopyMode = False
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Exit Sub

CleanFail:
    On Error Resume Next
    Set wsDash = ThisWorkbook.Sheets(DASH_SHEET)
    
    With wsDash
        .Range(AL_STATUS_CELL).Value = "FAILED"
        .Range(AL_LAST_REFRESH_CELL).Value = Now
        .Range(AL_ROWS_CELL).Value = 0
        .Range(AL_NOTES_CELL).Value = "Error: " & Err.Description
    End With
    
    MsgBox "Active Listing import failed: " & Err.Description, _
           vbCritical, "Active Listing Import"
    Resume CleanExit
End Sub

Private Sub CopyColumnWidths_AL(ByVal wsSource As Worksheet, ByVal wsDest As Worksheet)
    Dim lastCol As Long
    Dim c As Long
    
    lastCol = wsSource.UsedRange.Columns(wsSource.UsedRange.Columns.Count).Column
    
    For c = 1 To lastCol
        wsDest.Columns(c).ColumnWidth = wsSource.Columns(c).ColumnWidth
    Next c
End Sub
