Attribute VB_Name = "Module4"
Dim arg As String, x As Boolean, rowNum As Integer

Sub ListAllComplaints()
'Creates list from DB, ordered by complaint number
x = CheckDB
If Not x Then
    rowNum = Application.WorksheetFunction.Match("D*B*Loc*", Sheet1.Range("A:A"), 0) + 1
    If Range("A" & rowNum).Interior.ColorIndex = 3 Then
        Exit Sub 'error message thrown by checkDB above
    ElseIf Range("A" & rowNum).Interior.ColorIndex = 6 Then
        Exit Sub 'error message thrown by checkDB above
    Else
        MsgBox "An issue exists with the database and/or its directory. Repair this issue before proceeding."
    End If
    Exit Sub
End If
arg = "Complaints"
Call GenerateReport(arg, "text")
End Sub
Sub ListAllCustomerContacts()
'Creates list from DB, ordered by customer number
x = CheckDB
If Not x Then
    rowNum = Application.WorksheetFunction.Match("D*B*Loc*", Sheet1.Range("A:A"), 0) + 1
    If Range("A" & rowNum).Interior.ColorIndex = 3 Then
        Exit Sub 'error message thrown by checkDB above
    ElseIf Range("A" & rowNum).Interior.ColorIndex = 6 Then
        Exit Sub 'error message thrown by checkDB above
    Else
        MsgBox "An issue exists with the database and/or its directory. Repair this issue before proceeding."
    End If
    Exit Sub
End If
arg = "Customers"
Call GenerateReport(arg, "text")
End Sub
Sub ComplaintsByCat()
x = CheckDB
If Not x Then
    rowNum = Application.WorksheetFunction.Match("D*B*Loc*", Sheet1.Range("A:A"), 0) + 1
    If Range("A" & rowNum).Interior.ColorIndex = 3 Then
        Exit Sub 'error message thrown by checkDB above
    ElseIf Range("A" & rowNum).Interior.ColorIndex = 6 Then
        Exit Sub 'error message thrown by checkDB above
    Else
        MsgBox "An issue exists with the database and/or its directory. Repair this issue before proceeding."
    End If
    Exit Sub
End If
arg = "ComplaintsCat"
Call GenerateReport(arg, "graph")
End Sub
Sub ComplaintsByCust()
x = CheckDB
If Not x Then
    rowNum = Application.WorksheetFunction.Match("D*B*Loc*", Sheet1.Range("A:A"), 0) + 1
    If Range("A" & rowNum).Interior.ColorIndex = 3 Then
        Exit Sub 'error message thrown by checkDB above
    ElseIf Range("A" & rowNum).Interior.ColorIndex = 6 Then
        Exit Sub 'error message thrown by checkDB above
    Else
        MsgBox "An issue exists with the database and/or its directory. Repair this issue before proceeding."
    End If
    Exit Sub
End If
arg = "ComplaintsCust"
Call GenerateReport(arg, "graph")
End Sub
Sub ComplaintsBySup()
x = CheckDB
If Not x Then
    rowNum = Application.WorksheetFunction.Match("D*B*Loc*", Sheet1.Range("A:A"), 0) + 1
   If Range("A" & rowNum).Interior.ColorIndex = 3 Then
        Exit Sub 'error message thrown by checkDB above
    ElseIf Range("A" & rowNum).Interior.ColorIndex = 6 Then
        Exit Sub 'error message thrown by checkDB above
    Else
        MsgBox "An issue exists with the database and/or its directory. Repair this issue before proceeding."
    End If
    Exit Sub
End If
arg = "ComplaintsSup"
Call GenerateReport(arg, "graph")
End Sub
Sub ComplaintsByRC()
x = CheckDB
If Not x Then
    rowNum = Application.WorksheetFunction.Match("D*B*Loc*", Sheet1.Range("A:A"), 0) + 1
    If Range("A" & rowNum).Interior.ColorIndex = 3 Then
        Exit Sub 'error message thrown by checkDB above
    ElseIf Range("A" & rowNum).Interior.ColorIndex = 6 Then
        Exit Sub 'error message thrown by checkDB above
    Else
        MsgBox "An issue exists with the database and/or its directory. Repair this issue before proceeding."
    End If
    Exit Sub
End If
arg = "ComplaintsRC"
Call GenerateReport(arg, "graph")
End Sub

Function GenerateReport(repType As String, repClass As String)

If InStr(UCase(ThisWorkbook.Path), "ENGLANDT") = 0 Then
    MsgBox "Please use the correct workbook, not a copy."
    Exit Function
End If

Dim dbPath As String, cN As ADODB.Connection, recSet As ADODB.Recordset, recCol As ADODB.Recordset, tableName As String
Dim numRecs As Integer, i As Integer, custID As Integer, custName As String, custNum As Integer, compNo As String, fieldName As String

dbPath = Sheet1.Range("A" & Application.WorksheetFunction.Match("*D*b*Path*", Sheet1.Range("A:A"), 0) + 1).Value
'On Error GoTo errhandler

Set cN = New ADODB.Connection
cN.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & dbPath
cN.Open

If UCase(repClass) = "TEXT" Then
    Sheet5.Visible = xlSheetVisible
    Sheet5.Activate
    Sheet1.Visible = xlSheetHidden
Else
    Sheet6.Visible = xlSheetVisible
    Sheet6.Activate
    Sheet1.Visible = xlSheetHidden
End If

If repType = "Complaints" Then

    'list of complaints
    tableName = "ClaimInfo"
    
    Set recSet = New ADODB.Recordset
    recSet.Open tableName, cN, adOpenKeyset, adLockPessimistic, adCmdTable

    numRecs = recSet.RecordCount 'number of records/table rows

    If numRecs = 0 Then
        MsgBox "Issue locating customer records"
        Sheet1.Activate
        Sheet5.Visible = xlSheetHidden
        Exit Function
    End If

    'copy recordset into column 3
    Sheet5.Cells(3, 3).CopyFromRecordset recSet

    recSet.Close
    
    Columns("F:F").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    
    'get contact names
    tableName = "Contacts"

    Set recSet = New ADODB.Recordset
    recSet.Open tableName, cN, adOpenKeyset, adLockPessimistic, adCmdTable

    Sheet5.Activate
    Range("C3").Select
    Selection.End(xlDown).Select
    numRecs = ActiveCell.Row
    
    For i = 3 To numRecs
        custNum = Range("E" & i).Value
        recSet.Filter = "ID = '" & custNum & "'"
        Range("E" & i).Value = recSet.Fields("Contact")
        Range("F" & i).Value = recSet.Fields("Customer")
    Next i
    
    recSet.Close
    
    'get customer names
    tableName = "Customers"

    Set recSet = New ADODB.Recordset
    recSet.Open tableName, cN, adOpenKeyset, adLockPessimistic, adCmdTable
    
    For i = 3 To numRecs
        custID = Range("F" & i).Value
        recSet.Filter = "ID = '" & custID & "'"
        Range("F" & i).Value = recSet.Fields("Customer_Name")
    Next i
    
    recSet.Close
    Set recSet = Nothing
    cN.Close
    Set cN = Nothing
    
    
    Range("C2").Value = "Claim Number"
    Range("D2").Value = "Initiated By"
    Range("E2").Value = "Contact Name"
    Range("F2").Value = "Customer"
    Range("G2").Value = "Date Opened"
    Range("H2").Value = "RMA Number"
    Range("I2").Value = "Date Closed"
    Range("G:G").NumberFormat = "mm/dd/yy"
    Range("I:I").NumberFormat = "mm/dd/yy"
    Columns("C:C").HorizontalAlignment = xlCenter
    Columns("G:I").HorizontalAlignment = xlCenter
    Sheet5.Range("C2:P" & numRecs).Columns.AutoFit
    
ElseIf repType = "Customers" Then
    
    'list of contacts
    tableName = "Contacts"

    Set recSet = New ADODB.Recordset
    recSet.Open tableName, cN, adOpenKeyset, adLockPessimistic, adCmdTable

    numRecs = recSet.RecordCount 'number of records/table rows

    If numRecs = 0 Then
        MsgBox "Issue locating customer records"
        Sheet1.Activate
        Sheet5.Visible = xlSheetHidden
        Exit Function
    End If

    'copy recordset into column 3
    Sheet5.Cells(3, 3).CopyFromRecordset recSet

    recSet.Close
    
    'list of customers
    tableName = "Customers"

    Set recSet = New ADODB.Recordset
    recSet.Open tableName, cN, adOpenKeyset, adLockPessimistic, adCmdTable

    Sheet5.Activate
    Range("C3").Select
    Selection.End(xlDown).Select
    numRecs = ActiveCell.Row
    
    For i = 3 To numRecs
        custID = Range("E" & i).Value
        recSet.Filter = "ID = '" & custID & "'"
        Range("E" & i).Value = recSet.Fields("Customer_Name")
    Next i
    
    recSet.Close
    Set recSet = Nothing
    cN.Close
    Set cN = Nothing
    
    Range("C2").Value = "Record"
    Range("C:C").HorizontalAlignment = xlCenter
    Range("D2").Value = "Contact Name"
    Range("E2").Value = "Customer"
    Range("F2").Value = "Address"
    Range("G2").Value = "City"
    Range("H2").Value = "State"
    Range("I2").Value = "Zip Code"
    Range("J2").Value = "Country"
    Range("K2").Value = "Phone"
    Range("L2").Value = "Email"
    Sheet5.Range("C2:P" & numRecs).Columns.AutoFit
    
Else 'graphs
    
    tableName = "WarrantyLog"
    
    Set recSet = New ADODB.Recordset
    recSet.Open tableName, cN, adOpenKeyset, adLockPessimistic, adCmdTable
    
    numRecs = recSet.RecordCount 'number of records/table rows

    If numRecs = 0 Then
        MsgBox "Issue locating customer records"
        Sheet1.Activate
        Sheet6.Visible = xlSheetHidden
        Exit Function
    End If

    'copy recordset into column 11
    Sheet6.Cells(3, 11).CopyFromRecordset recSet

    recSet.Close

    If UCase(repType) = "COMPLAINTSCAT" Then
        fieldName = "Category"
        Range("C1").Value = "Complaints by Category"
        'Range("A3").Value = "Category"
    ElseIf UCase(repType) = "COMPLAINTSSUP" Then
        fieldName = "Supplier"
        Range("C1").Value = "Complaints by Supplier"
        'Range("A3").Value = "Supplier"
    ElseIf UCase(repType) = "COMPLAINTSRC" Then
        fieldName = "Root Cause Category"
        Range("C1").Value = "Complaints by Root Cause Category"
        'Range("A3").Value = "Root Cause Category"
    ElseIf UCase(repType) = "COMPLAINTSCUST" Then 'need to reference customer tables as well
        fieldName = "Customer"
        Range("C1").Value = "Complaints by Customer"
        'Range("A3").Value = "Customer"
        'add columns for contact/customr info...
        
        'use ClaimInfo to get Contact
        tableName = "ClaimInfo"
    
        Set recSet = New ADODB.Recordset
        recSet.Open tableName, cN, adOpenKeyset, adLockPessimistic, adCmdTable
    
        Sheet6.Activate
        Columns(15).EntireColumn.Insert '15 chosen randomly
        Range("O2").Value = "Contact ID"
        Range("K3").Select
        Selection.End(xlDown).Select
        numRecs = ActiveCell.Row
        
        For i = 3 To numRecs
            compNo = Range("L" & i).Value
            recSet.Filter = "Complaint_No = '" & compNo & "'"
            Range("O" & i).Value = recSet.Fields("CustomerContact")
        Next i
        recSet.Close
        
        'use Contacts to get Customer#
        tableName = "Contacts"
    
        Set recSet = New ADODB.Recordset
        recSet.Open tableName, cN, adOpenKeyset, adLockPessimistic, adCmdTable
    
        Sheet6.Activate
        Columns(15).EntireColumn.Insert '15 chosen randomly
        Range("O2").Value = "Customer ID"
        
        For i = 3 To numRecs
            custNum = Range("P" & i).Value
            recSet.Filter = "ID = '" & custNum & "'"
            Range("O" & i).Value = recSet.Fields("Customer")
        Next i
        recSet.Close
        
        'use Customers to get Customer Name
        tableName = "Customers"
    
        Set recSet = New ADODB.Recordset
        recSet.Open tableName, cN, adOpenKeyset, adLockPessimistic, adCmdTable
    
        Sheet6.Activate
        Columns(15).EntireColumn.Insert '15 chosen randomly
        Range("O2").Value = "Customer"
        
        For i = 3 To numRecs
            custID = Range("P" & i).Value
            recSet.Filter = "ID = '" & custID & "'"
            Range("O" & i).Value = recSet.Fields("Customer_Name")
        Next i
        recSet.Close
        
    End If
    
    cN.Close 'closed down here in case fieldname="Customer"
    'get column into 11
    Do While Cells(2, 11).Value <> fieldName
        Columns(11).EntireColumn.Delete
    Loop
    
    Do While Cells(2, 12).Value > 0
        Columns(12).EntireColumn.Delete
    Loop

    Range("K2").Clear
    Range("K3:K500").Copy
    Range("A4").PasteSpecial xlPasteValues
    Range("K3:K500").Clear
    Call UpdateDataRange

End If

Range("A1").Select

Exit Function
errhandler:
MsgBox "Error in GenerateReport sub"

End Function

Sub updatedata()
Call UpdateDataRange
End Sub
Function UpdateDataRange()

Dim pTable As PivotTable, newRange As String


newRange = Sheet6.Name & "!" & Range("$A$3:$A$750").Address(ReferenceStyle:=xlR1C1)

Set pTable = ActiveSheet.PivotTables("PivotTable13")

pTable.ChangePivotCache ActiveWorkbook. _
            PivotCaches.Create(SourceType:=xlDatabase, SourceData:=newRange, _
            Version:=xlPivotTableVersion15)
      
pTable.RefreshTable

End Function
Sub testpiv()
Call UpdateDataRange
End Sub

Sub BackFromTxtRep()

Call ClearTxtRepSheet
Call BackToMain
Application.ScreenUpdating = True

End Sub
Sub BackFromGraphRep()

Call ClearGraphRepSheet
Call BackToMain
Application.ScreenUpdating = True

End Sub

Function ClearTxtRepSheet()
Application.ScreenUpdating = False
Sheet5.Activate
Range("A1:Z500").Clear
Range("2:2").Font.Bold = True
Range("2:2").HorizontalAlignment = xlCenter

End Function

Function ClearGraphRepSheet()

Application.ScreenUpdating = False

Sheet6.Activate
Range("I2:Z500").Clear
Range("A5:A750").ClearContents
Range("A4").Value = "Test"
Range("2:2").Font.Bold = True
Range("2:2").HorizontalAlignment = xlCenter

Range("C1").Value = "Complaints Report"
Range("K2").Value = "Record"
Range("L2").Value = "Complaint"
Range("M2").Value = "Part No"
Range("N2").Value = "Serial No"
Range("O2").Value = "Mach Model"
Range("P2").Value = "Mach SN"
Range("Q2").Value = "Category"
Range("R2").Value = "Complaint"
Range("S2").Value = "Description"
Range("T2").Value = "Lot No"
Range("U2").Value = "Supplier"
Range("V2").Value = "Root Cause Category"
Range("W2").Value = "Root Cause"
Range("X2").Value = "SCAR"
Range("Y2").Value = "CAPA"

End Function
