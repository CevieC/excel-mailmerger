Option Explicit

' ==== CONFIG ====
Private Const ROOT_PATH As String = "C:\ROP_Letters"   ' <-- change this
Private Const FIELD_QUARTER As String = "Quarter"
Private Const FIELD_STATUS As String = "Active_Status"
Private Const FIELD_CHANNEL As String = "Channel_Folder"
Private Const FIELD_ADVISOR As String = "Producing_Advisor_Name"

Private Const EXCEL_SHEET_NAME As String = "ROP Letter"
Private Const EXCEL_PDF_HEADER As String = "PDF Path"

Public Sub LettersToPDF()
    Dim docMain As Document, docNew As Document
    Dim mm As MailMerge, ds As MailMergeDataSource
    Dim fso As Object, counter As Object
    Dim i As Long, key As String, idx As Long
    Dim q As String, s As String, ch As String, adv As String
    Dim folderPath As String, pdfName As String, pdfFullPath As String
    
    Dim xlApp As Object, wb As Object, wb2 As Object, ws As Object
    Dim pdfCol As Long, lastCol As Long, rowExcel As Long
    Dim fileName As String
    
    On Error GoTo ErrHandler
    
    Set docMain = ActiveDocument
    Set mm = docMain.MailMerge
    Set ds = mm.DataSource
    
    If ds.RecordCount = 0 Then
        MsgBox "No records in mail merge data source.", vbExclamation
        Exit Sub
    End If
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set counter = CreateObject("Scripting.Dictionary")
    
    ' --- Try to hook Excel for logging (optional) ---
    Set xlApp = Nothing: Set wb = Nothing: Set ws = Nothing
    On Error Resume Next
    Set xlApp = GetObject(, "Excel.Application")
    On Error GoTo ErrHandler
    
    If Not xlApp Is Nothing And ds.DataFiles.Count > 0 Then
        fileName = Dir(ds.DataFiles(1)) ' just "WorkbookA.xlsx"
        For Each wb2 In xlApp.Workbooks
            If StrComp(wb2.Name, fileName, vbTextCompare) = 0 Then
                Set wb = wb2
                Exit For
            End If
        Next wb2
        
        If Not wb Is Nothing Then
            On Error Resume Next
            Set ws = wb.Worksheets(EXCEL_SHEET_NAME)
            On Error GoTo ErrHandler
            
            If Not ws Is Nothing Then
                lastCol = ws.Cells(1, ws.Columns.Count).End(-4159).Column ' xlToLeft
                For pdfCol = 1 To lastCol
                    If Trim(CStr(ws.Cells(1, pdfCol).Value)) = EXCEL_PDF_HEADER Then Exit For
                Next pdfCol
                If pdfCol > lastCol Then
                    pdfCol = lastCol + 1
                    ws.Cells(1, pdfCol).Value = EXCEL_PDF_HEADER
                End If
            End If
        End If
    End If
    
    Application.ScreenUpdating = False
    
    With mm
        .Destination = wdSendToNewDocument
        .SuppressBlankLines = True
        
        For i = 1 To ds.RecordCount
            ds.ActiveRecord = i
            
            q = CleanText(FieldVal(ds, FIELD_QUARTER))
            s = CleanText(FieldVal(ds, FIELD_STATUS))
            ch = CleanText(FieldVal(ds, FIELD_CHANNEL))
            adv = CleanText(FieldVal(ds, FIELD_ADVISOR))
            
            If q = "" Then q = "Unknown Quarter"
            If s = "" Then s = "Unknown Status"
            If ch = "" Then ch = "Unknown Channel"
            If adv = "" Then adv = "Unknown Advisor"
            
            key = q & "|" & s & "|" & ch & "|" & adv
            If counter.Exists(key) Then
                idx = counter(key) + 1
            Else
                idx = 1
            End If
            counter(key) = idx
            
            folderPath = ROOT_PATH & "\" & _
                         SafePart(q) & "\" & SafePart(s) & "\" & SafePart(ch)
            EnsureFolder fso, folderPath
            
            pdfName = ch & " ROP Letter for " & q & " - " & adv & " " & idx & ".pdf"
            pdfName = SafeName(pdfName)
            pdfFullPath = folderPath & "\" & pdfName
            
            .DataSource.FirstRecord = i
            .DataSource.LastRecord = i
            .Execute Pause:=False
            
            Set docNew = ActiveDocument
            docNew.ExportAsFixedFormat pdfFullPath, wdExportFormatPDF, False
            docNew.Close False
            docMain.Activate
            
            ' log back to Excel if possible
            If Not ws Is Nothing Then
                rowExcel = i + 1          ' record 1 -> row 2
                ws.Cells(rowExcel, pdfCol).Value = pdfFullPath
            End If
        Next i
    End With
    
    If Not wb Is Nothing Then wb.Save
    
    Application.ScreenUpdating = True
    MsgBox "PDFs generated for " & ds.RecordCount & " records.", vbInformation
    Exit Sub
    
ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "Error: " & Err.Description, vbCritical
End Sub

' ----- helpers (short + shared) -----

Private Function FieldVal(ds As MailMergeDataSource, name As String) As String
    On Error Resume Next
    FieldVal = ds.DataFields(name).Value
    If Err.Number <> 0 Then FieldVal = "": Err.Clear
End Function

Private Function CleanText(t As String) As String
    t = Replace(t, "–", "-")
    t = Replace(t, "—", "-")
    t = Replace(t, vbCr, " ")
    t = Replace(t, vbLf, " ")
    t = Replace(t, vbTab, " ")
    Do While InStr(t, "  ") > 0
        t = Replace(t, "  ", " ")
    Loop
    CleanText = Trim(t)
End Function

Private Function SafeName(t As String) As String
    Dim bad: bad = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
    Dim x
    For Each x In bad
        t = Replace(t, CStr(x), " ")
    Next x
    t = CleanText(t)
    Do While Right$(t, 1) = "."
        t = Left$(t, Len(t) - 1)
    Loop
    If t = "" Then t = "ROP_Letter"
    SafeName = t
End Function

Private Function SafePart(t As String) As String
    t = SafeName(t)
    If t = "" Then t = "_"
    SafePart = t
End Function

Private Sub EnsureFolder(fso As Object, path As String)
    Dim parts() As String, cur As String, i As Long
    parts = Split(path, "\")
    cur = parts(0)
    For i = 1 To UBound(parts)
        cur = cur & "\" & parts(i)
        If Not fso.FolderExists(cur) Then fso.CreateFolder cur
    Next i
End Sub