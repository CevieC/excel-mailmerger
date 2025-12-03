Option Explicit

Private Sub Workbook_Open()
    Const AGTA_FOLDER As String = "C:\Path\To\Folder\"   
    Const DEST_SHEET As String = "AGTA"       
    Const DASH_SHEET As String = "Overview"             

    Const AGTA_FILE_NAME_CELL As String = "I13"
    Const AGTA_FILE_PATH_CELL As String = "I14"
    Const AGTA_LAST_REFRESH_CELL As String = "I15"
    Const AGTA_STATUS_CELL As String = "I16"
    Const AGTA_ROWS_CELL As String = "I17"
    Const AGTA_NOTES_CELL As String = "I18"
    
    Dim agtaName As String
    Dim agtaPath As String
    Dim wbSrc As Workbook
    Dim wsSrc As Worksheet
    Dim wsDest As Worksheet
    Dim wsDash As Worksheet
    Dim rowsLoaded As Long

    On Error GoTo CleanFail

    Set wsDash = ThisWorkbook.Sheets(DASH_SHEET)

    agtaName = "AGTA" & Format(Date, "MMYY") & ".xlsx"
    agtaPath = AGTA_FOLDER & agtaName

    With wsDash
        .Range(AGTA_FILE_NAME_CELL).Value = agtaName
        .Range(AGTA_FILE_PATH_CELL).Value = agtaPath
        .Range(AGTA_LAST_REFRESH_CELL).Value = Now
        .Range(AGTA_STATUS_CELL).Value = "FAILED"
        .Range(AGTA_ROWS_CELL).Value = 0
        .Range(AGTA_NOTES_CELL).Value = "Initialising AGTA refresh..."
    End With

    If Dir(agtaPath) = vbNullString Then
        With wsDash
            .Range(AGTA_STATUS_CELL).Value = "FAILED"
            .Range(AGTA_NOTES_CELL).Value = "File not found for current month."
        End With
        
        MsgBox "AGTA file not found:" & vbCrLf & agtaPath, _
               vbExclamation, "AGTA Refresh"
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    ' Open source workbook (read-only is fine)
    Set wbSrc = Workbooks.Open(agtaPath)
    Set wsSrc = wbSrc.Sheets(1)
    Set wsDest = ThisWorkbook.Sheets(DEST_SHEET)

    ' Clear destination sheet & copy filtered rows
    wsDest.Cells.Clear
    
    ' Copy header row first (assumes row 1 is headers)
    wsSrc.Rows(1).Copy wsDest.Rows(1)
    
    ' Copy column widths too (to keep formatting)
    CopyColumnWidths wsSrc, wsDest
    
    ' Filter and copy data rows where Col P = "FA" AND Col AS = "A"
    Dim srcLastRow As Long
    Dim srcRow As Long
    Dim destRow As Long
    Dim colP As Long
    Dim colAS As Long
    
    colP = 16   ' Column P
    colAS = 45  ' Column AS
    
    srcLastRow = wsSrc.Cells(wsSrc.Rows.Count, 1).End(xlUp).Row
    destRow = 2 ' Start pasting data from row 2 (after headers)
    
    For srcRow = 2 To srcLastRow ' Start from row 2 (skip headers)
        If wsSrc.Cells(srcRow, colP).Value = "FA" And _
           wsSrc.Cells(srcRow, colAS).Value = "A" Then
            
            wsSrc.Rows(srcRow).Copy wsDest.Rows(destRow)
            destRow = destRow + 1
        End If
    Next srcRow
    
    ' How many rows loaded (including header)?
    rowsLoaded = destRow - 1 ' Subtract 1 because destRow is now pointing to next empty row

    ' Close source without saving
    wbSrc.Close SaveChanges:=False

    ' Update dashboard for SUCCESS
    With wsDash
        .Range(AGTA_FILE_NAME_CELL).Value = agtaName
        .Range(AGTA_FILE_PATH_CELL).Value = agtaPath
        .Range(AGTA_LAST_REFRESH_CELL).Value = Now
        .Range(AGTA_STATUS_CELL).Value = "SUCCESS"
        .Range(AGTA_ROWS_CELL).Value = rowsLoaded
        .Range(AGTA_NOTES_CELL).Value = "OK"
    End With

CleanExit:
    Application.CutCopyMode = False
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Exit Sub

CleanFail:
    ' Any unexpected error -> log to dashboard as FAILED
    On Error Resume Next
    Set wsDash = ThisWorkbook.Sheets(DASH_SHEET)
    With wsDash
        .Range(AGTA_STATUS_CELL).Value = "FAILED"
        .Range(AGTA_LAST_REFRESH_CELL).Value = Now
        .Range(AGTA_ROWS_CELL).Value = 0
        .Range(AGTA_NOTES_CELL).Value = "Error: " & Err.Description
    End With
    On Error GoTo 0
    
    MsgBox "AGTA refresh failed: " & Err.Description, _
           vbCritical, "AGTA Refresh"
    Resume CleanExit
End Sub

Private Sub CopyColumnWidths(ByVal wsSource As Worksheet, ByVal wsDest As Worksheet)
    Dim lastCol As Long, c As Long
    lastCol = wsSource.UsedRange.Columns(wsSource.UsedRange.Columns.Count).Column
    
    For c = 1 To lastCol
        wsDest.Columns(c).ColumnWidth = wsSource.Columns(c).ColumnWidth
    Next c
End Sub
