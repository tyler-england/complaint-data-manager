Attribute VB_Name = "Module2"
Function CheckCapitalization(entryText As String, entryDesig As String) As String

Dim newText As String, firstLet As String, spaceEx As Boolean, secondLet As String, lastLet As String

Application.ScreenUpdating = False

On Error GoTo errhandler

CheckCapitalization = entryText
newText = Application.WorksheetFunction.Proper(entryText)

If entryText = newText Then
    Exit Function
End If

firstLet = Left(entryText, 1)

spaceEx = False
On Error Resume Next
    If InStr(entryText, " ") > 0 Then
        spaceEx = True
    End If
On Error GoTo errhandler

If spaceEx Then
    secondLet = Left(Right(entryText, Len(entryText) - InStr(entryText, " ")), 1) 'first letter, second word
Else
    secondLet = Right(Left(entryText, 2), 1) 'second letter, first word
End If

lastLet = Right(entryText, 1)

If spaceEx Then 'two words
    If firstLet = UCase(firstLet) And secondLet = UCase(secondLet) And lastLet = LCase(lastLet) Then
        Exit Function
    End If
Else 'one word
    If firstLet = UCase(firstLet) And secondLet = LCase(secondLet) And lastLet = LCase(lastLet) Then
        Exit Function
    End If
End If

If entryDesig = "Customer" And entryText = UCase(entryText) Then
    Exit Function
End If

ans = MsgBox("Would you like to change the " & entryDesig & " Name from...." & vbCrLf & vbCrLf & _
entryText & vbCrLf & vbCrLf & " to..." & vbCrLf & vbCrLf & newText, vbYesNo, "Potential Capitalization Error")

If ans = vbYes Then
    CheckCapitalization = newText
End If

Exit Function

errhandler:
MsgBox "Error in the CheckCapitalization function"

End Function

Function SetValidation(custID As Integer, custContRng As Range)

Dim sourceRange As Range, rowTop As Integer, rowBot As Integer, i As Integer

Application.ScreenUpdating = False

For i = 1 To 100
    If Sheet4.Range("D" & i).Value = custID Then
        rowTop = i
        Exit For
    End If
Next i

For i = rowTop To rowTop + 50
    If Sheet4.Range("D" & i).Value = Sheet4.Range("D" & rowTop).Value Then
        rowBot = i
    Else
        Exit For
    End If
Next i

If rowTop = rowBot Then 'only appears once
    Exit Function
End If

Set sourceRange = Sheet4.Range("C" & rowTop & ":C" & rowBot) 'sheet4 options for the given customer name

Sheet2.Unprotect
With custContRng.Validation
    .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
        Formula1:="='" & Sheet4.Name & "'!" & sourceRange.Address
    .ShowError = False
End With
Sheet2.Protect

End Function

Function PopulateCustomerDeets(custID As Integer, custContRng As Range)

Dim i As Integer, j As Integer, custRows() As Integer, contRows() As Integer, x As Integer
Dim custEntryExists As Boolean, contEntryExists As Boolean, infoRange As Range
Dim cAddress As String, cCity As String, cState As String, cZIP As String, cCountry As String

On Error GoTo errhandler

Application.ScreenUpdating = False

custEntryExists = False
contEntryExists = False

'Create array of rows for customer matches
j = 0
x = 0
ReDim Preserve custRows(x)
For i = 2 To 500
    If Sheet4.Range("D" & i).Value = custID Then 'customer is in the list
        custEntryExists = True
        ReDim Preserve custRows(0 To x) 'resize array of row numbers
        custRows(x) = i 'add row to array of row numbers
        x = x + 1
    ElseIf Sheet4.Range("D" & i).Value = 0 Then 'empty row
        j = j + 1
    End If
    
    If j > 1 Then 'multiple empty rows
        Exit For
    End If
Next i

If custEntryExists = False Then 'customer not in list
    Exit Function
End If

'Create array of rows for contact matches
j = 0
x = 0
ReDim Preserve contRows(x)

For i = 2 To 500
    
    If Sheet4.Range("C" & i).Value = custContRng.Value Then 'customer is in the list
        contEntryExists = True
        ReDim Preserve contRows(0 To x) 'resize array of row numbers
        contRows(x) = i 'add row to array of row numbers
        x = x + 1
    ElseIf Sheet4.Range("C" & i).Value = 0 Then 'empty row
        j = j + 1
    End If
    
    If j > 3 Then 'multiple empty rows
        Exit For
    End If
Next i

If contEntryExists = False Then
    Exit Function
End If

i = 0
j = 0
For Each custRow In custRows
    For Each contRow In contRows
        If custRow = contRow Then
            If i = 0 Then
                i = custRow
            Else
                j = custRow
            End If
        End If
    Next
    If j > 0 Then
        Exit For
    End If
Next

If i = 0 Then 'no matching rows between customers & contacts
    Exit Function
End If

If j > 0 Then 'customer/contact pair found for multiple entries
    MsgBox "The list of customers contains multiple entries for this customer/contact pair."
    Exit Function
End If

cAddress = Sheet4.Range("E" & i).Value
cCity = Sheet4.Range("F" & i).Value
cState = Sheet4.Range("G" & i).Value
cZIP = Sheet4.Range("H" & i).Value
cCountry = Sheet4.Range("I" & i).Value

Set infoRange = Sheet2.Range("B" & Application.WorksheetFunction.Match("Address*", Range("A:A"), 0))
infoRange.Value = cAddress
Set infoRange = infoRange.Offset(1, 0)
infoRange.Value = cCity
Set infoRange = infoRange.Offset(1, 0)
infoRange.Value = cState
Set infoRange = infoRange.Offset(1, 0)
infoRange.Value = cZIP
Set infoRange = infoRange.Offset(1, 0)
infoRange.Value = cCountry

Exit Function

errhandler:
    MsgBox "Error in PopulateCustomerDeets function"
    
End Function

Function UpdateCustomerDropdown()

Dim i As Integer, x As Integer, customerList() As String, numRows As Integer

Application.ScreenUpdating = False

On Error GoTo errhandler
Exit Sub
'this sub was used to update the dropdown but now the data
'are all pulled directly from the Access database
Sheet4.Activate
Sheet4.Range("A2").Select
Selection.End(xlDown).Select
Range("A2:A" & ActiveCell.Row).Select
Selection.Clear

Sheet4.Range("B2").Select
Selection.End(xlDown).Select
numRows = ActiveCell.Row

ReDim customerList(0)
customerList(0) = Sheet4.Range("B2").Value

x = 1
For i = 3 To numRows

    If Sheet4.Range("B" & i).Value <> Sheet4.Range("B" & i - 1).Value Then
        ReDim Preserve customerList(x)
        customerList(x) = Sheet4.Range("B" & i).Value
        x = x + 1
    End If

Next i

For i = 0 To UBound(customerList)
    
    Sheet4.Range("A" & i + 2).Value = customerList(i)
    
Next i

Sheet4.Range("A1").Select
Exit Function

errhandler:
MsgBox "Error in UpdateCustomerDropdown function"
End Function

Function AddToWarrantyTable(i As Integer) As Boolean

Dim dbPath As String, tableName As String, cN As ADODB.Connection, recSet As ADODB.Recordset
Dim j As Integer

On Error GoTo errhandler

Application.ScreenUpdating = False

AddToWarrantyTable = False

dbPath = Sheet1.Range("A" & Application.WorksheetFunction.Match("Full*D*B*", Sheet1.Range("A:A"), 0) + 1).Value
tableName = "WarrantyLog"

Set cN = New ADODB.Connection
cN.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & dbPath
cN.Open

Set recSet = New ADODB.Recordset
recSet.Open tableName, cN, adOpenKeyset, adLockPessimistic, adCmdTable

For j = 6 To i

    recSet.AddNew
    recSet.Fields("Part_No") = Cells(j, Application.WorksheetFunction.Match("Part*Num*", Sheet2.Range("5:5"), 0)).Value
    recSet.Fields("Serial_No") = Cells(j, Application.WorksheetFunction.Match("Part*SN*", Sheet2.Range("5:5"), 0)).Value
    recSet.Fields("Complaint_No") = Range("B" & Application.WorksheetFunction.Match("Complaint*", Sheet2.Range("A:A"), 0)).Value
    recSet.Fields("Machine_SN") = Cells(j, Application.WorksheetFunction.Match("Machine*SN*", Sheet2.Range("5:5"), 0)).Value
    recSet.Fields("Complaint_Cat") = Cells(j, Application.WorksheetFunction.Match("Complaint*Cat*", Sheet2.Range("5:5"), 0)).Value
    recSet.Fields("Complaint") = Cells(j, Application.WorksheetFunction.Match("Complaint", Sheet2.Range("5:5"), 0)).Value
    recSet.Fields("Item_Description") = Cells(j, Application.WorksheetFunction.Match("*Description*", Sheet2.Range("5:5"), 0)).Value
    recSet.Fields("Machine_Model") = Cells(j, Application.WorksheetFunction.Match("Machine*Model*", Sheet2.Range("5:5"), 0)).Value
    recSet.Fields("Lot_No") = Cells(j, Application.WorksheetFunction.Match("Lot*N*", Sheet2.Range("5:5"), 0)).Value
    If Cells(j, Application.WorksheetFunction.Match("*Supplier*", Sheet2.Range("5:5"), 0)).Value <> "Suppliers" Then
        recSet.Fields("Supplier") = Cells(j, Application.WorksheetFunction.Match("*Supplier*", Sheet2.Range("5:5"), 0)).Value
    End If
    recSet.Update
Next j

recSet.Close
Set recSet = Nothing
cN.Close
Set cN = Nothing

AddToWarrantyTable = True

Exit Function

errhandler:
MsgBox "Error in AddToWarrantyTable function"

End Function

Function AddToComplaintTable() As Boolean

Dim dbPath As String, tableName As String, cN As ADODB.Connection, recSet As ADODB.Recordset
Dim claimNo As String, psaName As String, custNum As Integer, dateOpen As Date, rmaNo As String
Dim custName As String, contName As String, custID As Integer

On Error GoTo errhandler

Application.ScreenUpdating = False
AddToComplaintTable = False

claimNo = Sheet2.Range("B" & Application.WorksheetFunction.Match("Complaint*", Sheet2.Range("A:A"), 0)).Value
psaName = Sheet2.Range("B" & Application.WorksheetFunction.Match("Your*", Sheet2.Range("A:A"), 0)).Value
dateOpen = Sheet2.Range("B" & Application.WorksheetFunction.Match("*Date*", Sheet2.Range("A:A"), 0)).Value
rmaNo = Sheet2.Range("B" & Application.WorksheetFunction.Match("RMA*", Sheet2.Range("A:A"), 0)).Value
dbPath = Sheet1.Range("A" & Application.WorksheetFunction.Match("Full*D*B*", Sheet1.Range("A:A"), 0) + 1).Value

tableName = "Customers" 'Need to get customer ID for customer
custName = Sheet2.Range("B" & Application.WorksheetFunction.Match("Customer*", Sheet2.Range("A:A"), 0)).Value
contName = Sheet2.Range("B" & Application.WorksheetFunction.Match("Contact*", Sheet2.Range("A:A"), 0)).Value

Set cN = New ADODB.Connection
cN.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & dbPath
cN.Open

Set recSet = New ADODB.Recordset
recSet.Open tableName, cN, adOpenKeyset, adLockPessimistic, adCmdTable

recSet.Filter = "Customer_Name = '" & custName & "'"

If recSet.RecordCount <> 1 Then
    MsgBox "Error occurred identifying Customer ID. Database not updated."
    Exit Function
End If

custID = recSet.Fields("ID") 'now that we have ID num, can proceed
recSet.Close

'get record for specific contact person
tableName = "Contacts"

Set recSet = New ADODB.Recordset
recSet.Open tableName, cN, adOpenKeyset, adLockPessimistic, adCmdTable

recSet.Filter = "Customer = '" & custID & "' AND Contact = '" & contName & "'"

If recSet.RecordCount <> 1 Then
    MsgBox "Error occurred identifying Customer ID. Database not updated."
    Exit Function
End If

custNum = recSet.Fields("ID") 'now that we have ID num, can proceed
recSet.Close

tableName = "ClaimInfo"

Set recSet = New ADODB.Recordset
recSet.Open tableName, cN, adOpenKeyset, adLockPessimistic, adCmdTable

recSet.Filter = "Complaint_No = '" & claimNo & "'"

If recSet.RecordCount <> 0 Then 'record with that claim no exists
    MsgBox "There is an issue with the Complaint Number entered. Perhaps it is invalid, " & _
            "or perhaps it is already in the database. Check the Complaint Number and " & _
            "retry. The database has not been updated."
    Exit Function
End If

recSet.AddNew
recSet.Fields("Complaint_No") = claimNo
recSet.Fields("Initiated_By") = psaName
recSet.Fields("CustomerContact") = custNum
recSet.Fields("Date_Opened") = dateOpen
recSet.Fields("RMA_No") = rmaNo
recSet.Update

recSet.Close
Set recSet = Nothing
cN.Close
Set cN = Nothing

AddToComplaintTable = True

Exit Function

errhandler:
MsgBox "Error in AddToComplaintTable function"

End Function


Function UpdateCustomerInfo() As Boolean

Dim dbPath As String, cName As String, cContact As String, cAddress As String, cPhone As String
Dim cCity As String, cState As String, cZIP As String, cCountry As String, cEmail As String
Dim tableName As String, cN As ADODB.Connection, recSet As ADODB.Recordset, custID As Integer

On Error GoTo errhandler

Application.ScreenUpdating = False
UpdateCustomerInfo = False

cName = Sheet2.Range("B" & Application.WorksheetFunction.Match("Customer*", Sheet2.Range("A:A"), 0)).Value
cContact = Sheet2.Range("B" & Application.WorksheetFunction.Match("Contact*", Sheet2.Range("A:A"), 0)).Value
cPhone = Sheet2.Range("B" & Application.WorksheetFunction.Match("Phone*", Sheet2.Range("A:A"), 0)).Value
cEmail = Sheet2.Range("B" & Application.WorksheetFunction.Match("Email*", Sheet2.Range("A:A"), 0)).Value
cAddress = Sheet2.Range("B" & Application.WorksheetFunction.Match("Address*", Sheet2.Range("A:A"), 0)).Value
cCity = Sheet2.Range("B" & Application.WorksheetFunction.Match("City*", Sheet2.Range("A:A"), 0)).Value
cState = Sheet2.Range("B" & Application.WorksheetFunction.Match("State*", Sheet2.Range("A:A"), 0)).Value
cZIP = Sheet2.Range("B" & Application.WorksheetFunction.Match("Zip*", Sheet2.Range("A:A"), 0)).Value
cCountry = Sheet2.Range("B" & Application.WorksheetFunction.Match("Country*", Sheet2.Range("A:A"), 0)).Value

dbPath = Sheet1.Range("A" & Application.WorksheetFunction.Match("Full*D*B*", Sheet1.Range("A:A"), 0) + 1).Value
tableName = "Customers"

Set cN = New ADODB.Connection
cN.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & dbPath
cN.Open

'get customer ID
Set recSet = New ADODB.Recordset
recSet.Open tableName, cN, adOpenKeyset, adLockPessimistic, adCmdTable

recSet.Filter = "Customer_Name = '" & cName & "'"

If recSet.RecordCount = 0 Then
    recSet.AddNew
    recSet.Fields("Customer_Name") = cName
End If
custID = recSet.Fields("ID")

recSet.Close

'check for contacts
tableName = "Contacts"
Set recSet = New ADODB.Recordset
recSet.Open tableName, cN, adOpenKeyset, adLockPessimistic, adCmdTable

recSet.Filter = "Customer = '" & custID & "' AND Contact = '" & cContact & "'"

If recSet.RecordCount = 0 Then 'no matching  record exists --> create it
    recSet.AddNew
End If

recSet.Fields("Customer") = custID
recSet.Fields("Contact") = cContact
recSet.Fields("Address") = cAddress
recSet.Fields("City") = cCity
recSet.Fields("State") = cState
recSet.Fields("ZIP") = cZIP
recSet.Fields("Country") = cCountry
recSet.Fields("Phone") = cPhone
recSet.Fields("Email") = cEmail
recSet.Update

recSet.Close
Set recSet = Nothing
cN.Close
Set cN = Nothing

UpdateCustomerInfo = True

Exit Function

errhandler:

MsgBox "Error in UpdateCustomerInfo function"
End Function

Sub CreateBackupDB()

Dim wrkgObj As Object, dbPath As String, backupPath As String, destPath As String, fileName As String, folderName As String

On Error GoTo errhandler

Set wrkgObj = CreateObject("Scripting.FileSystemObject")

backupPath = "\\PSACLW02\ProjData\EnglandT\Misc\Backups\WarrantyDB\"
dbPath = Sheet1.Range("A" & Application.WorksheetFunction.Match("Full*D*B*", Sheet1.Range("A:A"), 0) + 1).Value

If Right(backupPath, 1) <> "\" Then
    backupPath = backupPath & "\"
End If

fileName = Dir(dbPath)
folderName = Dir(backupPath, vbDirectory)

If folderName = "" Then
    MkDir backupPath
    folderName = Dir(backupPath, vbDirectory)
End If

destPath = backupPath & Format(Now, "yyyymmdd_hhnnss") & "-" & fileName

If fileName > "" And folderName > "" Then

    wrkgObj.CopyFile dbPath, destPath
    'using this because filecopy will not work if the db is open by someone
    
End If

Set wrkgObj = Nothing

Exit Sub

errhandler:
If Application.userName = "England, Tyler (PSA-CLW)" Then
    MsgBox "Unable to create database backup"
End If

End Sub

Sub ManualClearSheet()
Call ClearSheet(30)
End Sub
