Option Explicit

Private Sub Workbook_Open()
    Const AGTA_FOLDER As String = "C:\Path\To\Folder\"   ' Default AGTA folder (auto), you can change this
    Const DEST_SHEET As String = "AGTA"                  ' Where data will be imported
    Const DASH_SHEET As String = "Overview"              ' Dashboard sheet name

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

    ' Build default expected AGTA file name (AGTAMMYY.xlsx)
    agtaName = "AGTA" & Format(Date, "MMYY") & ".xlsx"
    agtaPath = AGTA_FOLDER & agtaName

    ' Initialise dashboard as FAILED (defensive default)
    With wsDash
        .Range(AGTA_FILE_NAME_CELL).Value = agtaName
        .Range(AGTA_FILE_PATH_CELL).Value = agtaPath
        .Range(AGTA_LAST_REFRESH_CELL).Value = Now
        .Range(AGTA_STATUS_CELL).Value = "FAILED"
        .Range(AGTA_ROWS_CELL).Value = 0
        .Range(AGTA_NOTES_CELL).Value = "Initialising AGTA refresh..."
    End With

    ' === 1) Try default AGTA file first ===
    If Dir(agtaPath) = vbNullString Then
        ' Default file not found → prompt user to pick AGTA file manually
        wsDash.Range(AGTA_NOTES_CELL).Value = "Default AGTA file not found. Please select a file manually."

        If Not SelectAGTAFile(agtaPath, agtaName) Then
            ' User cancelled the file picker
            With wsDash
                .Range(AGTA_STATUS_CELL).Value = "FAILED"
                .Range(AGTA_ROWS_CELL).Value = 0
                .Range(AGTA_NOTES_CELL).Value = "User cancelled AGTA file selection."
            End With

            MsgBox "AGTA refresh cancelled. No file selected.", _
                   vbExclamation, "AGTA Refresh"
            GoTo CleanExit
        End If

        ' Update dashboard with the chosen file details
        With wsDash
            .Range(AGTA_FILE_NAME_CELL).Value = agtaName
            .Range(AGTA_FILE_PATH_CELL).Value = agtaPath
            .Range(AGTA_LAST_REFRESH_CELL).Value = Now
            .Range(AGTA_STATUS_CELL).Value = "PENDING"
            .Range(AGTA_ROWS_CELL).Value = 0
            .Range(AGTA_NOTES_CELL).Value = "Importing AGTA from manually selected file..."
        End With
    Else
        ' Default file exists → note that we will use it
        wsDash.Range(AGTA_NOTES_CELL).Value = "Importing AGTA from default monthly file..."
    End If

    ' === 2) Proceed to import using agtaPath (either default or user-selected) ===
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    ' Open source workbook (read-only is fine)
    Set wbSrc = Workbooks.Open(agtaPath)
    Set wsSrc = wbSrc.Sheets(1)
    Set wsDest = ThisWorkbook.Sheets(DEST_SHEET)

    ' Clear destination sheet & copy everything
    wsDest.Cells.Clear
    wsSrc.UsedRange.Copy wsDest.Range("A1")

    ' Copy column widths too (to keep formatting)
    CopyColumnWidths wsSrc, wsDest

    ' How many rows loaded?
    rowsLoaded = wsDest.UsedRange.Rows.Count

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

' Helper: copy column widths from source to destination
Private Sub CopyColumnWidths(ByVal wsSource As Worksheet, ByVal wsDest As Worksheet)
    Dim lastCol As Long, c As Long
    lastCol = wsSource.UsedRange.Columns(wsSource.UsedRange.Columns.Count).Column
    
    For c = 1 To lastCol
        wsDest.Columns(c).ColumnWidth = wsSource.Columns(c).ColumnWidth
    Next c
End Sub

' Helper: open File Explorer so user can manually pick an AGTA file
Private Function SelectAGTAFile(ByRef agtaPath As String, ByRef agtaName As String) As Boolean
    Dim fd As FileDialog
    Dim result As Integer

    SelectAGTAFile = False   ' default

    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "Select AGTA file to import"
        ' ▼ Placeholder start folder – change this later to your preferred path
        .InitialFileName = "C:\Placeholder\AGTA\Folder\"    ' e.g. "C:\Users\You\AGTA Files\"
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xlsx;*.xlsm;*.xlsb;*.xls"
        .AllowMultiSelect = False

        result = .Show
        If result <> -1 Then
            ' User pressed Cancel / closed dialog
            Exit Function
        End If

        agtaPath = .SelectedItems(1)
        agtaName = Dir(agtaPath)
    End With

    SelectAGTAFile = True
End Function
