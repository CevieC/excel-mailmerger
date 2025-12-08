Option Explicit

' Build / refresh the ROP staging sheet from the raw ROP data
Public Sub BuildROPStaging()
    Const SRC_SHEET As String = "Data"
    Const STAGING_SHEET As String = "ROP_Letter"
    
    Dim wsSrc As Worksheet
    Dim wsStg As Worksheet
    Dim lastRow As Long
    Dim r As Long
    
    Dim master As Object        ' Scripting.Dictionary (key: Policy Owner ID, item: info dict)
    Dim info As Object          ' per-client dictionary
    Dim oldPolicies As Object   ' dict of old policies (key: Policy No, item: Policy Name)
    Dim newPolicies As Object   ' dict of new policies (key: Policy No, item: Policy Name)
    
    Dim policyOwnerID As String
    Dim policyOwnerName As String
    Dim insuredName As String
    Dim insuredID As String
    Dim advisorName As String
    Dim advisorCode As String
    Dim oldPolNo As String
    Dim oldPolName As String
    Dim newPolNo As String
    Dim newPolName As String
    
    Dim outRow As Long
    Dim k As Variant
    Dim pol As Variant
    
    On Error GoTo ErrHandler
    
    ' Get source sheet
    On Error Resume Next
    Set wsSrc = ThisWorkbook.Worksheets(SRC_SHEET)
    On Error GoTo ErrHandler
    
    If wsSrc Is Nothing Then
        MsgBox "Source sheet '" & SRC_SHEET & "' not found. Please check the name in the code.", _
               vbExclamation, "Build ROP Staging"
        Exit Sub
    End If
    
    ' Determine last row based on Policy Owner ID column (D)
    lastRow = wsSrc.Cells(wsSrc.Rows.Count, "D").End(xlUp).Row
    If lastRow < 2 Then
        MsgBox "No data found in sheet '" & SRC_SHEET & "'.", vbInformation, "Build ROP Staging"
        Exit Sub
    End If
    
    ' Get / create staging sheet
    On Error Resume Next
    Set wsStg = ThisWorkbook.Worksheets(STAGING_SHEET)
    On Error GoTo ErrHandler
    
    If wsStg Is Nothing Then
        Set wsStg = ThisWorkbook.Worksheets.Add(After:=wsSrc)
        wsStg.Name = STAGING_SHEET
    End If
    
    ' Clear only data rows (row 2 onwards) in columns A to J, preserving headers and columns after J
    Dim lastStgRow As Long
    lastStgRow = wsStg.Cells(wsStg.Rows.Count, "A").End(xlUp).Row
    If lastStgRow >= 2 Then
        wsStg.Range("A2:J" & lastStgRow).Clear
    End If
    
    ' Build master dictionary grouped by Policy Owner ID
    Set master = CreateObject("Scripting.Dictionary")
    
    Application.ScreenUpdating = False
    
    For r = 2 To lastRow
        policyOwnerID = Trim(CStr(wsSrc.Cells(r, "D").Value))   ' Policy Owner ID (unique identifier)
        
        ' Only process rows with Policy Owner ID
        If policyOwnerID <> "" Then
            ' Create per-client info dict if new
            If Not master.Exists(policyOwnerID) Then
                Set info = CreateObject("Scripting.Dictionary")
                
                ' Store basic info (from first occurrence of this Policy Owner ID)
                info("PolicyOwnerName") = Trim(CStr(wsSrc.Cells(r, "C").Value))
                info("PolicyOwnerID") = policyOwnerID
                info("InsuredName") = Trim(CStr(wsSrc.Cells(r, "A").Value))
                info("InsuredID") = Trim(CStr(wsSrc.Cells(r, "B").Value))
                info("AdvisorName") = Trim(CStr(wsSrc.Cells(r, "V").Value))
                info("AdvisorCode") = Trim(CStr(wsSrc.Cells(r, "S").Value))
                
                ' Create dictionaries for policies
                Set oldPolicies = CreateObject("Scripting.Dictionary")
                Set newPolicies = CreateObject("Scripting.Dictionary")
                
                ' Store policy dictionaries in info dict
                Set info("OldPolicies") = oldPolicies
                Set info("NewPolicies") = newPolicies
                
                master.Add policyOwnerID, info
            Else
                Set info = master(policyOwnerID)
                Set oldPolicies = info("OldPolicies")
                Set newPolicies = info("NewPolicies")
            End If
            
            ' Collect OLD policy for this row
            oldPolNo = Trim(CStr(wsSrc.Cells(r, "AM").Value))   ' OLD Policy No
            oldPolName = Trim(CStr(wsSrc.Cells(r, "AV").Value))  ' OLD Policy Name
            If oldPolNo <> "" Then
                If Not oldPolicies.Exists(oldPolNo) Then
                    oldPolicies.Add oldPolNo, oldPolName
                End If
            End If
            
            ' Collect NEW policy for this row
            newPolNo = Trim(CStr(wsSrc.Cells(r, "F").Value))    ' NEW Policy No
            newPolName = Trim(CStr(wsSrc.Cells(r, "O").Value))  ' NEW Policy Name
            If newPolNo <> "" Then
                If Not newPolicies.Exists(newPolNo) Then
                    newPolicies.Add newPolNo, newPolName
                End If
            End If
        End If
    Next r
    
    ' Output to staging sheet
    outRow = 2
    
    For Each k In master.Keys
        Set info = master(k)
        
        ' Retrieve policy dictionaries
        On Error Resume Next
        Set oldPolicies = info("OldPolicies")
        Set newPolicies = info("NewPolicies")
        On Error GoTo ErrHandler
        
        ' Write basic info
        wsStg.Cells(outRow, "A").Value = info("PolicyOwnerName")
        wsStg.Cells(outRow, "B").Value = info("PolicyOwnerID")
        wsStg.Cells(outRow, "C").Value = info("InsuredName")
        wsStg.Cells(outRow, "D").Value = info("InsuredID")
        wsStg.Cells(outRow, "E").Value = info("AdvisorName")
        wsStg.Cells(outRow, "F").Value = info("AdvisorCode")
        
        ' Build OLD Policies block
        wsStg.Cells(outRow, "G").Value = BuildPoliciesBlock(oldPolicies)
        wsStg.Cells(outRow, "H").Value = oldPolicies.Count
        
        ' Build NEW Policies block
        wsStg.Cells(outRow, "I").Value = BuildPoliciesBlock(newPolicies)
        wsStg.Cells(outRow, "J").Value = newPolicies.Count
        
        outRow = outRow + 1
    Next k
    
    ' Format staging sheet
    With wsStg
        .Columns("A:J").EntireColumn.AutoFit
        .Columns("G:I").WrapText = True
        .Rows("1:1").Font.Bold = True
    End With
    
    Application.ScreenUpdating = True
    
    MsgBox "ROP staging build complete. " & (outRow - 2) & " unique clients processed.", _
           vbInformation, "Build ROP Staging"
    
    Exit Sub
    
ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "Error while building ROP staging: " & Err.Description, vbCritical, "Build ROP Staging"
End Sub

' Helper: builds the multi-line block for policies (format: <Policy Number> <Policy Name>)
Private Function BuildPoliciesBlock(ByVal dict As Object) As String
    Dim result As String
    Dim pol As Variant
    Dim polName As String
    Dim firstItem As Boolean
    
    If dict Is Nothing Then
        BuildPoliciesBlock = vbNullString
        Exit Function
    End If
    
    If dict.Count = 0 Then
        BuildPoliciesBlock = vbNullString
        Exit Function
    End If
    
    result = ""
    firstItem = True
    
    For Each pol In dict.Keys
        If Not firstItem Then
            result = result & vbLf
        End If
        
        On Error Resume Next
        polName = CStr(dict(pol))
        If Err.Number <> 0 Then
            polName = ""
            Err.Clear
        End If
        On Error GoTo 0
        
        ' Format: <Policy Number> <Policy Name>
        result = result & CStr(pol)
        If Len(polName) > 0 Then
            result = result & " " & polName
        End If
        
        firstItem = False
    Next pol
    
    BuildPoliciesBlock = result
End Function

