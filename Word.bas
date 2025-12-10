Option Explicit

' === USER SETTINGS ===
Private Const ROOT_PATH As String = "C:\ROP_Letters"   ' <-- CHANGE THIS TO YOUR DIRECTORY

' Mail merge field names (these MUST match your ROP Letter headers)
Private Const FIELD_QUARTER As String = "Quarter"
Private Const FIELD_STATUS_FOLDER As String = "Active_Status"
Private Const FIELD_CHANNEL_FOLDER As String = "Channel_Folder"
Private Const FIELD_ADVISOR_NAME As String = "Producing_Advisor_Name"

' Excel sheet / header settings
Private Const EXCEL_SHEET_NAME As String = "ROP Letter"
Private Const EXCEL_PDF_PATH_HEADER As String = "PDF Path"

' ============================================================
'   MAIN MACRO – GENERATE PDFS AND WRITE BACK TO EXCEL
' ============================================================

Public Sub GenerateROPLettersToPDF()

    Dim docMain As Document
    Dim docNew As Document
    Dim mm As MailMerge
    Dim ds As MailMergeDataSource
    
    Dim fso As Object
    Dim advisorLetterCount As Object
    
    Dim i As Long
    Dim quarterVal As String, statusFolder As String
    Dim channelFolder As String, advisorName As String
    Dim advisorKey As String, letterIndex As Long
    
    Dim targetFolder As String
    Dim pdfFileName As String
    Dim pdfFullPath As String
    
    ' Excel objects
    Dim xlApp As Object, wb As Object, ws As Object
    Dim pdfCol As Long, lastCol As Long, rowInExcel As Long
    Dim canLogToExcel As Boolean
    
    On Error GoTo ErrHandler
    
    Set docMain = ActiveDocument
    Set mm = docMain.MailMerge
    Set ds = mm.DataSource
    
    If ds.RecordCount = 0 Then
        MsgBox "No records found in the mail merge data source.", vbExclamation
        Exit Sub
    End If
    
    ' Create filesystem + advisor count dict
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set advisorLetterCount = CreateObject("Scripting.Dictionary")
    
    canLogToExcel = False
    
    On Error Resume Next
    Set xlApp = GetObject(Class:="Excel.Application")
    On Error GoTo ErrHandler
    
    If Not xlApp Is Nothing Then
        
        Set wb = xlApp.Workbooks(ds.Name)     ' workbook name must match mail merge data source file name
        If Not wb Is Nothing Then
            
            Set ws = wb.Worksheets(EXCEL_SHEET_NAME)
            If Not ws Is Nothing Then
                
                ' Find or create PDF Path column
                pdfCol = 0
                lastCol = ws.Cells(1, ws.Columns.Count).End(-4159).Column
                
                Dim c As Long
                For c = 1 To lastCol
                    If Trim(CStr(ws.Cells(1, c).Value)) = EXCEL_PDF_PATH_HEADER Then
                        pdfCol = c
                        Exit For
                    End If
                Next c
                
                If pdfCol = 0 Then
                    pdfCol = lastCol + 1
                    ws.Cells(1, pdfCol).Value = EXCEL_PDF_PATH_HEADER
                End If
                
                canLogToExcel = True
            End If
        End If
    End If
    
    Application.ScreenUpdating = False
    
    ' ----------------------------
    ' PROCESS EACH RECORD
    ' ----------------------------
    
    With mm
        .Destination = wdSendToNewDocument
        .SuppressBlankLines = True
        
        For i = 1 To ds.RecordCount
        
            ds.ActiveRecord = i
            
            ' Read merge fields
            quarterVal = CleanText(GetDataFieldValue(ds, FIELD_QUARTER))
            statusFolder = CleanText(GetDataFieldValue(ds, FIELD_STATUS_FOLDER))
            channelFolder = CleanText(GetDataFieldValue(ds, FIELD_CHANNEL_FOLDER))
            advisorName = CleanText(GetDataFieldValue(ds, FIELD_ADVISOR_NAME))
            
            If quarterVal = "" Then quarterVal = "Unknown Quarter"
            If statusFolder = "" Then statusFolder = "Unknown Status"
            If channelFolder = "" Then channelFolder = "Unknown Channel"
            If advisorName = "" Then advisorName = "Unknown Advisor"
            
            ' Counting logic for file numbering
            advisorKey = quarterVal & "|" & statusFolder & "|" & channelFolder & "|" & advisorName
            
            If advisorLetterCount.Exists(advisorKey) Then
                letterIndex = advisorLetterCount(advisorKey) + 1
            Else
                letterIndex = 1
            End If
            advisorLetterCount(advisorKey) = letterIndex
            
            ' Build folder path
            targetFolder = ROOT_PATH & "\" & _
                           SanitizePathComponent(quarterVal) & "\" & _
                           SanitizePathComponent(statusFolder) & "\" & _
                           SanitizePathComponent(channelFolder)
            
            EnsureFolderExists fso, targetFolder
            
            ' Build PDF filename
            pdfFileName = channelFolder & " ROP Letter for " & quarterVal & " - " & advisorName & " " & letterIndex & ".pdf"
            pdfFileName = SanitizeFileName(pdfFileName)
            
            pdfFullPath = targetFolder & "\" & pdfFileName
            
            ' Merge this single record → new document
            .DataSource.FirstRecord = i
            .DataSource.LastRecord = i
            .Execute Pause:=False
            
            Set docNew = ActiveDocument
            
            ' Export PDF
            docNew.ExportAsFixedFormat _
                OutputFileName:=pdfFullPath, _
                ExportFormat:=wdExportFormatPDF, _
                OpenAfterExport:=False
            
            docNew.Close SaveChanges:=False
            
            ' Write into Excel (row = i + 1)
            If canLogToExcel Then
                rowInExcel = i + 1
                ws.Cells(rowInExcel, pdfCol).Value = pdfFullPath
            End If
            
            docMain.Activate
        Next i
    End With
    
    Application.ScreenUpdating = True
    
    If canLogToExcel Then wb.Save
    
    MsgBox "PDF letters generated successfully for " & ds.RecordCount & " records.", vbInformation
    Exit Sub

' Error handler
ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "Error generating letters:" & vbCrLf & Err.Description, vbCritical
End Sub


' ============================================================
'                    HELPER FUNCTIONS
' ============================================================

Private Function GetDataFieldValue(ds As MailMergeDataSource, fieldName As String) As String
    On Error GoTo ErrHandler
    GetDataFieldValue = ds.DataFields(fieldName).Value
    Exit Function
ErrHandler:
    GetDataFieldValue = ""
End Function

Private Function CleanText(txt As String) As String
    txt = Replace(txt, "–", "-")
    txt = Replace(txt, "—", "-")
    txt = Replace(txt, vbCr, " ")
    txt = Replace(txt, vbLf, " ")
    txt = Replace(txt, vbTab, " ")
    Do While InStr(txt, "  ") > 0
        txt = Replace(txt, "  ", " ")
    Loop
    CleanText = Trim(txt)
End Function

Private Function SanitizeFileName(fileName As String) As String
    Dim badChars: badChars = Array("\", "/", ":", "*", "?", """", "<", ">","|")
    Dim ch
    For Each ch In badChars
        fileName = Replace(fileName, ch, " ")
    Next ch
    fileName = CleanText(fileName)
    Do While Right(fileName, 1) = "."
        fileName = Left(fileName, Len(fileName) - 1)
    Loop
    If Trim(fileName) = "" Then fileName = "ROP_Letter"
    SanitizeFileName = Trim(fileName)
End Function

Private Function SanitizePathComponent(part As String) As String
    part = SanitizeFileName(part)
    If part = "" Then part = "_"
    SanitizePathComponent = part
End Function

Private Sub EnsureFolderExists(fso As Object, folderPath As String)
    Dim parts() As String: parts = Split(folderPath, "\")
    Dim current As String: current = parts(0)
    
    Dim i As Long
    For i = 1 To UBound(parts)
        current = current & "\" & parts(i)
        If Not fso.FolderExists(current) Then fso.CreateFolder current
    Next i
End Sub