Option Explicit

' Header names in ROP Letter sheet
Private Const HDR_SHEET      As String = "ROP Letter"
Private Const HDR_PDF_PATH   As String = "PDF Path"
Private Const HDR_EMAIL_TO   As String = "Email To"
Private Const HDR_EMAIL_CC   As String = "Email CC"
Private Const HDR_QUARTER    As String = "Quarter"
Private Const HDR_ADVISOR    As String = "Producing Advisor Name"

Public Sub CreateROPLetterEmailDrafts()
    Dim ws As Worksheet
    Dim lastRow As Long, r As Long
    Dim colPdf As Long, colTo As Long, colCc As Long
    Dim colQ As Long, colAdv As Long
    
    Dim pdfPath As String
    Dim emailTo As String, emailCc As String
    Dim quarterVal As String, advisorName As String
    
    Dim olApp As Object, olMail As Object
    Dim subjectText As String, bodyText As String
    
    On Error GoTo ErrHandler
    
    Set ws = ThisWorkbook.Worksheets(HDR_SHEET)
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If lastRow < 2 Then
        MsgBox "No data rows found in '" & HDR_SHEET & "'.", vbExclamation
        Exit Sub
    End If
    
    ' Get column indices by header text in row 1
    colPdf = GetHeaderCol(ws, HDR_PDF_PATH)
    colTo = GetHeaderCol(ws, HDR_EMAIL_TO)
    colCc = GetHeaderCol(ws, HDR_EMAIL_CC)
    colQ = GetHeaderCol(ws, HDR_QUARTER)
    colAdv = GetHeaderCol(ws, HDR_ADVISOR)
    
    If colPdf * colTo * colCc = 0 Then
        MsgBox "Missing one or more required headers: " & _
               HDR_PDF_PATH & ", " & HDR_EMAIL_TO & ", " & HDR_EMAIL_CC, vbCritical
        Exit Sub
    End If
    
    ' Get or create Outlook
    On Error Resume Next
    Set olApp = GetObject(, "Outlook.Application")
    If olApp Is Nothing Then Set olApp = CreateObject("Outlook.Application")
    On Error GoTo ErrHandler
    
    Application.ScreenUpdating = False
    
    For r = 2 To lastRow
        pdfPath = Trim$(CStr(ws.Cells(r, colPdf).Value))
        emailTo = Trim$(CStr(ws.Cells(r, colTo).Value))
        emailCc = Trim$(CStr(ws.Cells(r, colCc).Value))
        
        If colQ > 0 Then quarterVal = CleanText(CStr(ws.Cells(r, colQ).Value)) Else quarterVal = ""
        If colAdv > 0 Then advisorName = CleanText(CStr(ws.Cells(r, colAdv).Value)) Else advisorName = ""
        
        If pdfPath = "" Then GoTo NextRow
        If Dir$(pdfPath) = "" Then
            Debug.Print "PDF not found for row " & r & ": " & pdfPath
            GoTo NextRow
        End If
        
        ' Build subject/body (customize as you like)
        If quarterVal = "" Then quarterVal = "this period"
        If advisorName = "" Then advisorName = "Advisor"
        
        subjectText = "ROP Letter for " & quarterVal & " - " & advisorName
        bodyText = "Dear " & advisorName & "," & vbCrLf & vbCrLf & _
                   "Please find attached your ROP letter." & vbCrLf & vbCrLf & _
                   "Regards," & vbCrLf & _
                   "<Your Name Here>"
        
        If Not olApp Is Nothing Then
            Set olMail = olApp.CreateItem(0) ' MailItem
            With olMail
                .To = emailTo
                .CC = emailCc
                .Subject = subjectText
                .Body = bodyText
                .Attachments.Add pdfPath
                .Display   ' change to .Save if you don't want popups
            End With
        End If
        
NextRow:
    Next r
    
    Application.ScreenUpdating = True
    MsgBox "Email drafts created for rows 2 to " & lastRow & ".", vbInformation
    Exit Sub
    
ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "Error creating email drafts: " & Err.Description, vbCritical
End Sub

' === Helpers ===

Private Function GetHeaderCol(ws As Worksheet, headerText As String) As Long
    Dim lastCol As Long, c As Long
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    For c = 1 To lastCol
        If Trim$(CStr(ws.Cells(1, c).Value)) = headerText Then
            GetHeaderCol = c
            Exit Function
        End If
    Next c
    GetHeaderCol = 0
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
    CleanText = Trim$(t)
End Function
