Option Explicit

' === CONFIGURABLE SETTINGS ===
Private Const ROOT_PATH As String = "C:\ROP_Letters"  ' TODO: change this to your real root folder

' These must match the column headers in the ROP Letter sheet / MailMerge data source
Private Const FIELD_QUARTER As String = "Quarter"
Private Const FIELD_STATUS_FOLDER As String = "Active Status"        ' e.g. "Active" / "Terminated"
Private Const FIELD_CHANNEL_FOLDER As String = "Channel Folder"      ' e.g. "Direct", "Agency", "FA", etc.
Private Const FIELD_ADVISOR_NAME As String = "Producing Advisor Name"

' Main macro
Public Sub LettersToPDF()
    Dim docMain As Document
    Dim docNew As Document
    Dim mm As MailMerge
    Dim ds As MailMergeDataSource
    
    Dim fso As Object
    Dim advisorLetterCount As Object    ' Dictionary to track numbering per advisor
    
    Dim i As Long
    Dim quarterVal As String
    Dim statusFolder As String
    Dim channelFolder As String
    Dim advisorName As String
    
    Dim advisorKey As String
    Dim letterIndex As Long
    
    Dim targetFolder As String
    Dim pdfFileName As String
    Dim pdfFullPath As String
    
    On Error GoTo ErrHandler
    
    Set docMain = ActiveDocument
    Set mm = docMain.MailMerge
    Set ds = mm.DataSource
    
    If ds.RecordCount = 0 Then
        MsgBox "No records found in the mail merge data source.", vbExclamation, "Generate ROP Letters"
        Exit Sub
    End If
    
    ' File system + advisor counters
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set advisorLetterCount = CreateObject("Scripting.Dictionary")
    
    Application.ScreenUpdating = False
    
    With mm
        .Destination = wdSendToNewDocument
        .SuppressBlankLines = True
        
        ' Loop through each record in the data source
        For i = 1 To ds.RecordCount
            ds.ActiveRecord = i
            
            ' --- Read values from current record ---
            quarterVal = Trim(GetDataFieldValue(ds, FIELD_QUARTER))
            statusFolder = Trim(GetDataFieldValue(ds, FIELD_STATUS_FOLDER))        ' "Active" / "Terminated"
            channelFolder = Trim(GetDataFieldValue(ds, FIELD_CHANNEL_FOLDER))      ' "Direct", "Agency", "FA", etc.
            advisorName = Trim(GetDataFieldValue(ds, FIELD_ADVISOR_NAME))
            
            ' Fallbacks to avoid empty pieces
            If quarterVal = "" Then quarterVal = "Unknown Quarter"
            If statusFolder = "" Then statusFolder = "Unknown Status"
            If channelFolder = "" Then channelFolder = "Unknown Channel"
            If advisorName = "" Then advisorName = "Unknown Advisor"
            
            ' --- Build key for numbering per advisor within this grouping ---
            advisorKey = quarterVal & "|" & statusFolder & "|" & channelFolder & "|" & advisorName
            
            If advisorLetterCount.Exists(advisorKey) Then
                letterIndex = CLng(advisorLetterCount(advisorKey)) + 1
            Else
                letterIndex = 1
            End If
            advisorLetterCount(advisorKey) = letterIndex
            
            ' --- Build folder path ---
            ' <ROOT>\<Quarter>\<Active Status>\<Channel Folder>\<Producing Advisor Name>\
            targetFolder = ROOT_PATH & "\" & _
                           SanitizePathComponent(quarterVal) & "\" & _
                           SanitizePathComponent(statusFolder) & "\" & _
                           SanitizePathComponent(channelFolder) & "\" & _
                           SanitizePathComponent(advisorName)
            
            EnsureFolderExists fso, targetFolder
            
            ' --- Build PDF file name ---
            ' <Channel Folder> ROP Letter for <Quarter> - <Producing Advisor Name> <n>.pdf
            pdfFileName = channelFolder & " ROP Letter for " & quarterVal & " - " & advisorName & " " & CStr(letterIndex) & ".pdf"
            pdfFileName = SanitizeFileName(pdfFileName)
            
            pdfFullPath = targetFolder & "\" & pdfFileName
            
            ' --- Execute merge for this single record ---
            .DataSource.FirstRecord = i
            .DataSource.LastRecord = i
            .Execute Pause:=False   ' creates a new document
            
            Set docNew = ActiveDocument
            
            ' --- Export the new document as PDF ---
            docNew.ExportAsFixedFormat _
                OutputFileName:=pdfFullPath, _
                ExportFormat:=wdExportFormatPDF, _
                OpenAfterExport:=False, _
                OptimizeFor:=wdExportOptimizeForPrint, _
                Range:=wdExportAllDocument, _
                Item:=wdExportDocumentContent, _
                IncludeDocProps:=True, _
                KeepIRM:=True, _
                CreateBookmarks:=wdExportCreateNoBookmarks, _
                DocStructureTags:=True, _
                BitmapMissingFonts:=True, _
                UseISO19005_1:=False
            
            ' Close the merged document without saving
            docNew.Close SaveChanges:=False
            
            ' Reactivate main template document
            docMain.Activate
        Next i
    End With
    
    Application.ScreenUpdating = True
    
    MsgBox "ROP letters generated successfully for " & ds.RecordCount & " records.", _
           vbInformation, "Generate ROP Letters"
    
    Exit Sub
    
ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "Error during ROP letter generation:" & vbCrLf & Err.Description, _
           vbCritical, "Generate ROP Letters"
End Sub

' Safely get a field value (returns "" if field not found)
Private Function GetDataFieldValue(ByVal ds As MailMergeDataSource, ByVal fieldName As String) As String
    On Error GoTo ErrHandler
    GetDataFieldValue = ds.DataFields(fieldName).Value
    Exit Function
ErrHandler:
    GetDataFieldValue = ""
End Function

' Remove invalid filename characters
Private Function SanitizeFileName(ByVal fileName As String) As String
    Dim badChars As Variant
    Dim ch As Variant
    
    badChars = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
    
    For Each ch In badChars
        fileName = Replace(fileName, CStr(ch), " ")
    Next ch
    
    fileName = Trim(fileName)
    SanitizeFileName = fileName
End Function

' Sanitize folder path component (reuse filename rules)
Private Function SanitizePathComponent(ByVal part As String) As String
    part = SanitizeFileName(part)
    If part = "" Then part = "_"
    SanitizePathComponent = part
End Function

' Ensure folder (and its parents) exist using FileSystemObject
Private Sub EnsureFolderExists(ByVal fso As Object, ByVal folderPath As String)
    Dim parts() As String
    Dim currentPath As String
    Dim i As Long
    
    If fso.FolderExists(folderPath) Then Exit Sub
    
    parts = Split(folderPath, "\")
    
    currentPath = parts(0)
    For i = 1 To UBound(parts)
        currentPath = currentPath & "\" & parts(i)
        If Not fso.FolderExists(currentPath) Then
            fso.CreateFolder currentPath
        End If
    Next i
End Sub
