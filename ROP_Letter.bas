Option Explicit

' Build / refresh the ROP staging sheet from the raw ROP data
Public Sub BuildROPStaging()
    Const SRC_SHEET As String = "Data"       
    Const STAGING_SHEET As String = "ROP_Letter"
    
    Dim wsSrc As Worksheet
    Dim wsStg As Worksheet
    Dim lastRow As Long
    Dim r As Long
    
    Dim master As Object        ' Scripting.Dictionary (key: Agent||NRIC, item: info dict)
    Dim info As Object          ' per-key dictionary
    Dim oldDict As Object       ' dict of old policies for this key
    Dim newDict As Object       ' dict of new policies for this key
    
    Dim key As String
    Dim agentName As String
    Dim nric As String
    Dim laName As String
    Dim oldPol As String, oldDesc As String
    Dim newPol As String, newDesc As String
    
    Dim outRow As Long
    Dim k As Variant
    
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
    
    ' Determine last row based on Agent Name column (BF)
    lastRow = wsSrc.Cells(wsSrc.Rows.Count, "BF").End(xlUp).Row
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
    
    ' Clear staging sheet and set headers
    With wsStg
        .Cells.Clear
        .Range("A1").Value = "Agent Name"
        .Range("B1").Value = "Life Assured NRIC"
        .Range("C1").Value = "Life Assured Name"
        .Range("D1").Value = "Count New Policies"
        .Range("E1").Value = "Count Old Policies"
        .Range("F1").Value = "New Policies Block"
        .Range("G1").Value = "Old Policies Block"
    End With
    
    ' Build master dictionary
    Set master = CreateObject("Scripting.Dictionary")
    
    Application.ScreenUpdating = False
    
    For r = 2 To lastRow
        agentName = Trim(wsSrc.Cells(r, "BF").Value)   ' Agent Name
        nric = Trim(wsSrc.Cells(r, "B").Value)         ' NRIC
        laName = Trim(wsSrc.Cells(r, "A").Value)       ' Life Assured Name
        
        ' Only process rows with agent + NRIC
        If agentName <> "" And nric <> "" Then
            key = agentName & "||" & nric
            
            ' Create per-key info dict if new
            If Not master.Exists(key) Then
                Set info = CreateObject("Scripting.Dictionary")
                info("Agent") = agentName
                info("NRIC") = nric
                info("Name") = laName
                
                Set oldDict = CreateObject("Scripting.Dictionary")
                Set newDict = CreateObject("Scripting.Dictionary")
                info("OldDict") = oldDict
                info("NewDict") = newDict
                
                master.Add key, info
            Else
                Set info = master(key)
                Set oldDict = info("OldDict")
                Set newDict = info("NewDict")
            End If
            
            ' Old policy for this row
            oldPol = Trim(wsSrc.Cells(r, "F").Value)   ' Old Policy No
            oldDesc = Trim(wsSrc.Cells(r, "O").Value)  ' Old Policy Desc
            If oldPol <> "" Then
                If Not oldDict.Exists(oldPol) Then
                    oldDict.Add oldPol, oldDesc
                End If
            End If
            
            ' New policy for this row
            newPol = Trim(wsSrc.Cells(r, "AM").Value)   ' New Policy No
            newDesc = Trim(wsSrc.Cells(r, "AV").Value)  ' New Policy Desc
            If newPol <> "" Then
                If Not newDict.Exists(newPol) Then
                    newDict.Add newPol, newDesc
                End If
            End If
        End If
    Next r
    
    ' Output to staging sheet
    outRow = 2
    
    For Each k In master.Keys
        Set info = master(k)
        Set oldDict = info("OldDict")
        Set newDict = info("NewDict")
        
        wsStg.Cells(outRow, "A").Value = info("Agent")
        wsStg.Cells(outRow, "B").Value = info("NRIC")
        wsStg.Cells(outRow, "C").Value = info("Name")
        
        wsStg.Cells(outRow, "D").Value = newDict.Count
        wsStg.Cells(outRow, "E").Value = oldDict.Count
        
        wsStg.Cells(outRow, "F").Value = BuildPoliciesBlock("New Policies:", newDict)
        wsStg.Cells(outRow, "G").Value = BuildPoliciesBlock("Old Policies:", oldDict)
        
        outRow = outRow + 1
    Next k
    
    ' Format staging sheet
    With wsStg
        .Columns("A:G").EntireColumn.AutoFit
        .Columns("F:G").WrapText = True
        .Rows("1:1").Font.Bold = True
    End With
    
    Application.ScreenUpdating = True
    
    MsgBox "ROP staging build complete. " & (outRow - 2) & " agent/NRIC combinations processed.", _
           vbInformation, "Build ROP Staging"
    
    Exit Sub
    
ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "Error while building ROP staging: " & Err.Description, vbCritical, "Build ROP Staging"
End Sub

' Helper: builds the multi-line block for new/old policies
Private Function BuildPoliciesBlock(ByVal header As String, ByVal dict As Object) As String
    Dim result As String
    Dim pol As Variant
    Dim desc As String
    Dim firstItem As Boolean
    
    If dict Is Nothing Then
        BuildPoliciesBlock = vbNullString
        Exit Function
    End If
    
    If dict.Count = 0 Then
        BuildPoliciesBlock = vbNullString
        Exit Function
    End If
    
    result = header & vbLf
    firstItem = True
    
    For Each pol In dict.Keys
        If Not firstItem Then
            result = result & vbLf
        End If
        
        On Error Resume Next
        desc = CStr(dict(pol))
        If Err.Number <> 0 Then
            desc = ""
            Err.Clear
        End If
        On Error GoTo 0
        
        result = result & CStr(pol)
        If Len(desc) > 0 Then
            result = result & " (" & desc & ")"
        End If
        
        firstItem = False
    Next pol
    
    BuildPoliciesBlock = result
End Function
