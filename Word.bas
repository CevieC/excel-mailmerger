Option Explicit

Sub MergeToPDFs()

    Dim mainDoc As Document
    Dim mm As MailMerge
    Dim i As Long, totalRecs As Long
    
    Dim fileName As String
    Dim idValue As String
    Dim statusValue As String
    Dim channelValue As String
    
    Dim rootFolder As String
    Dim statusFolder As String
    Dim channelFolder As String
    Dim targetFolder As String
    
    ' 🔹 SET THESE: field names from your mail merge / Excel
    Const KEY_FIELD As String = "Name"        ' used in filename (e.g. PolicyNo / Name)
    Const STATUS_FIELD As String = "Status"   ' col L (A / T)
    Const CHANNEL_FIELD As String = "Channel" ' col N
    
    ' 🔹 SET THIS: root folder for all agent files (no need to end with \)
    rootFolder = "C:\Agent_Files"
    
    Set mainDoc = ActiveDocument
    Set mm = mainDoc.MailMerge
    
    If mm.State <> wdMainAndDataSource Then
        MsgBox "This document is not a mail merge main document.", vbExclamation
        Exit Sub
    End If
    
    totalRecs = mm.DataSource.RecordCount
    If totalRecs < 1 Then
        MsgBox "No records found in the mail merge data source.", vbExclamation
        Exit Sub
    End If
    
    ' Ensure root folder exists
    rootFolder = AddTrailingBackslash(rootFolder)
    EnsureFolderExists rootFolder
    
    Application.ScreenUpdating = False
    
    For i = 1 To totalRecs
        
        With mm.DataSource
            .FirstRecord = i
            .LastRecord = i
            
            On Error Resume Next
            idValue = .DataFields(KEY_FIELD).Value
            statusValue = .DataFields(STATUS_FIELD).Value
            channelValue = .DataFields(CHANNEL_FIELD).Value
            On Error GoTo 0
        End With
        
        idValue = Trim$(idValue)
        statusValue = UCase$(Trim$(statusValue))
        channelValue = Trim$(channelValue)
        
        ' Fallback filename if key field is blank
        If idValue = "" Then
            idValue = "Record_" & CStr(i)
        End If
        
        ' Clean filename
        fileName = CleanForFileName(idValue)
        
        ' Normalise status (A / T / UnknownStatus)
        Select Case statusValue
            Case "A", "T"
                ' ok
            Case Else
                statusValue = "UnknownStatus"
        End Select
        
        ' Fallback channel if empty
        If channelValue = "" Then
            channelValue = "NoChannel"
        End If
        
        ' Clean channel name for folder
        channelValue = CleanForFileName(channelValue)
        
        ' Build folder paths
        statusFolder = rootFolder & statusValue & "\"
        channelFolder = statusFolder & channelValue & "\"
        
        ' Ensure folders exist
        EnsureFolderExists statusFolder
        EnsureFolderExists channelFolder
        
        targetFolder = channelFolder
        
        ' Perform merge of this single record to new doc
        mm.Destination = wdSendToNewDocument
        mm.Execute Pause:=False
        
        ' Export the newly created document as PDF
        With ActiveDocument
            .ExportAsFixedFormat _
                OutputFileName:=targetFolder & fileName & ".pdf", _
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
            
            .Close SaveChanges:=False   ' close merged doc
        End With
        
        mainDoc.Activate
        
    Next i
    
    Application.ScreenUpdating = True
    
    MsgBox "Done! " & totalRecs & " PDF(s) created in:" & vbCrLf & rootFolder, vbInformation

End Sub

'==== Helpers ====

Private Function AddTrailingBackslash(ByVal folderPath As String) As String
    If Len(folderPath) = 0 Then
        AddTrailingBackslash = ""
    ElseIf Right$(folderPath, 1) = "\" Then
        AddTrailingBackslash = folderPath
    Else
        AddTrailingBackslash = folderPath & "\"
    End If
End Function

Private Sub EnsureFolderExists(ByVal folderPath As String)
    Dim fso As Object
    If Len(folderPath) = 0 Then Exit Sub
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    On Error Resume Next
    If Not fso.FolderExists(folderPath) Then
        fso.CreateFolder folderPath
    End If
    On Error GoTo 0
End Sub

Private Function CleanForFileName(ByVal s As String) As String
    ' Remove characters that are illegal in Windows file/folder names
    s = Replace$(s, "/", "-")
    s = Replace$(s, "\", "-")
    s = Replace$(s, ":", " ")
    s = Replace$(s, "*", " ")
    s = Replace$(s, "?", " ")
    s = Replace$(s, """", "'")
    s = Replace$(s, "<", " ")
    s = Replace$(s, ">", " ")
    s = Replace$(s, "|", " ")
    CleanForFileName = Trim$(s)
End Function
