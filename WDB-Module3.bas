Attribute VB_Name = "Module3"
Function CheckDB() As Boolean

Dim dbFolder As String, dbFileName As String, readWrite As Boolean, testFolName As String
Dim dbFolRow As Integer, dbNameRow As String, filePresent As Boolean, colIndex As Integer, colNum As Integer

On Error GoTo errhandler

CheckDB = False

Application.ScreenUpdating = False

'Check DB folder situation
Sheet1.Activate
dbFolRow = Application.WorksheetFunction.Match("D*B*Location*", Sheet1.Range("A:A"), 0) 'row with db loc header
dbFolder = Sheet1.Range("A" & dbFolRow + 1).Value
If Right(dbFolder, 1) <> "\" Then
    dbFolder = dbFolder & "\"
End If

dbNameRow = Application.WorksheetFunction.Match("*D*B*Name*", Sheet1.Range("A:A"), 0)  'row with db filename header
dbFileName = Sheet1.Range("A" & dbNameRow + 1).Value

'Check that DB is in specified directory
dbFileName = Dir(dbFolder & dbFileName & "*")

filePresent = False
If dbFileName > "" Then 'a file was found. correct file? maybe
    Do While Right(dbFileName, 6) <> ".accdb"
        dbFileName = Dir()
        If dbFileName = "" Then
            dbFileName = ".accdb" 'exit loop
        End If
    Loop

    If dbFileName <> ".accdb" Then 'filename is not empty
        filePresent = True
    End If
End If

colIndex = 0
If Not (filePresent) Then 'no DB file with that name in that folder
    CheckDB = False
    MsgBox "Error locating database. Check to make sure database " & _
            "location and database name are both correct."
    colIndex = 3 'red
End If

colNum = Application.WorksheetFunction.Match("*Report*", Sheet1.Range("1:1"), 0) - 2
Sheet1.Range(Cells(dbFolRow, 1), Cells(dbFolRow + 1, colNum)).Interior.ColorIndex = colIndex
Sheet1.Range(Cells(dbNameRow, 1), Cells(dbNameRow + 1, colNum)).Interior.ColorIndex = colIndex

If Not filePresent Then
    Exit Function
End If

On Error Resume Next
    testFolName = "Test Folder (To Be Deleted)"
    MkDir dbFolder & testFolName 'try making a folder
    
    If Err.Number <> 0 Then 'folder wasn't made
        readWrite = False
    Else 'folder was made
        readWrite = True
    End If
    
    RmDir dbFolder & testFolName 'delete it (if it doesn't exist, error will resume next)
On Error GoTo errhandler

colIndex = 0
CheckDB = True
If Not (readWrite) Then 'DB can't be modified
    CheckDB = False
    MsgBox "You don't seem to have read/write access to the location " & _
            "where the database is stored. Contact your IT administrator " & _
            "in order to obtain access."
    colIndex = 6 'yellow
End If

Sheet1.Range(Cells(dbFolRow + 1, 1), Cells(dbFolRow + 1, colNum)).Interior.ColorIndex = colIndex

Exit Function

errhandler:
MsgBox "Error in CheckDB function"

End Function

Function CheckForClaim(ccNo As String) As Boolean

Dim dbPath As String, tableName As String, cN As ADODB.Connection, recSet As ADODB.Recordset

On Error GoTo errhandler

Application.ScreenUpdating = False

CheckForClaim = False

dbPath = Sheet1.Range("A" & Application.WorksheetFunction.Match("Full*D*B*", Sheet1.Range("A:A"), 0) + 1).Value
tableName = "ClaimInfo"

Set cN = New ADODB.Connection
cN.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & dbPath
cN.Open

Set recSet = New ADODB.Recordset
recSet.Open tableName, cN, adOpenKeyset, adLockPessimistic, adCmdTable

recSet.Filter = "Complaint_No = '" & ccNo & "'"

If recSet.RecordCount = 1 Then
    CheckForClaim = True
End If

recSet.Close
Set recSet = Nothing
cN.Close
Set cN = Nothing

Exit Function

errhandler:
MsgBox "Error in CheckForClaim function"

End Function

Function ShowAllClaimInfo(ccNo As String)

Dim dbPath As String, tableName As String, cN As ADODB.Connection, recSet As ADODB.Recordset, custNum As Integer
Dim psaName As String, cName As String, cContact As String, cAddress As String, cCity As String
Dim cState As String, cZIP As String, cCountry As String, cDateOpen As Date, rmaNo As String
Dim cDateClose As Date, numRecs As Integer, i As Integer, j As Integer

On Error GoTo errhandler

Application.ScreenUpdating = False

dbPath = Sheet1.Range("A" & Application.WorksheetFunction.Match("Full*D*B*", Sheet1.Range("A:A"), 0) + 1).Value
tableName = "ClaimInfo"

Set cN = New ADODB.Connection
cN.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & dbPath
cN.Open

Set recSet = New ADODB.Recordset
recSet.Open tableName, cN, adOpenKeyset, adLockPessimistic, adCmdTable

On Error Resume Next 'blank DB values cause errors
    recSet.Filter = "Complaint_No = '" & ccNo & "'"
    custNum = recSet.Fields("CustomerContact") 'used to get customer info
    psaName = recSet.Fields("Initiated_By")
    cDateOpen = recSet.Fields("Date_Opened")
    rmaNo = recSet.Fields("RMA_No")
    cDateClose = recSet.Fields("Date_Closed")
On Error GoTo errhandler

'get customer info
tableName = "Contacts"

Set recSet = New ADODB.Recordset
recSet.Open tableName, cN, adOpenKeyset, adLockPessimistic, adCmdTable

recSet.Filter = "ID = '" & custNum & "'"

If recSet.RecordCount <> 1 Then
    MsgBox "Error finding customer info for this claim."
End If

On Error Resume Next 'blank DB values cause errors
    custNum = recSet.Fields("Customer")
    cContact = recSet.Fields("Contact")
    cAddress = recSet.Fields("Address")
    cCity = recSet.Fields("City")
    cState = recSet.Fields("State")
    cZIP = recSet.Fields("ZIP")
    cCountry = recSet.Fields("Country")
On Error GoTo errhandler
recSet.Close

'get customer name from custnum
tableName = "Customers"

Set recSet = New ADODB.Recordset
recSet.Open tableName, cN, adOpenKeyset, adLockPessimistic, adCmdTable

recSet.Filter = "ID = '" & custNum & "'"

If recSet.RecordCount <> 1 Then
    MsgBox "Error finding customer info for this claim."
End If

On Error Resume Next 'blank DB values cause errors
    cName = recSet.Fields("Customer_Name")
On Error GoTo errhandler

'fill in claim basics
Sheet3.Activate
Range("B" & Application.WorksheetFunction.Match("Complaint*", Range("A:A"), 0)).Value = ccNo
Range("B" & Application.WorksheetFunction.Match("Quality*", Range("A:A"), 0)).Value = psaName
Range("B" & Application.WorksheetFunction.Match("Customer*", Range("A:A"), 0)).Value = cName
Range("B" & Application.WorksheetFunction.Match("Contact*", Range("A:A"), 0)).Value = cContact
Range("B" & Application.WorksheetFunction.Match("Address*", Range("A:A"), 0)).Value = cAddress
If cCity > "" Then
    Range("B" & Application.WorksheetFunction.Match("City*", Range("A:A"), 0)).Value = cCity
    Range("D" & Application.WorksheetFunction.Match("City*", Range("A:A"), 0)).Value = False
Else
    Range("B" & Application.WorksheetFunction.Match("City*", Range("A:A"), 0)).Value = ""
    Range("D" & Application.WorksheetFunction.Match("City*", Range("A:A"), 0)).Value = True
End If
If cState > "" Then
    Range("B" & Application.WorksheetFunction.Match("State*", Range("A:A"), 0)).Value = cState
    Range("D" & Application.WorksheetFunction.Match("State*", Range("A:A"), 0)).Value = False
Else
    Range("B" & Application.WorksheetFunction.Match("State*", Range("A:A"), 0)).Value = ""
    Range("D" & Application.WorksheetFunction.Match("State*", Range("A:A"), 0)).Value = True
End If
If cZIP > "" Then
    Range("B" & Application.WorksheetFunction.Match("ZIP*", Range("A:A"), 0)).Value = cZIP
    Range("D" & Application.WorksheetFunction.Match("ZIP*", Range("A:A"), 0)).Value = False
Else
    Range("B" & Application.WorksheetFunction.Match("ZIP*", Range("A:A"), 0)).Value = ""
    Range("D" & Application.WorksheetFunction.Match("ZIP*", Range("A:A"), 0)).Value = True
End If
Range("B" & Application.WorksheetFunction.Match("Country*", Range("A:A"), 0)).Value = cCountry
If cDateOpen > 0 Then
    Range("B" & Application.WorksheetFunction.Match("*Open*", Range("A:A"), 0)).Value = cDateOpen
Else
    Range("B" & Application.WorksheetFunction.Match("*Open*", Range("A:A"), 0)).Value = ""
End If
If cDateClose > 0 Then
    Range("B" & Application.WorksheetFunction.Match("*Close*", Range("A:A"), 0)).Value = cDateClose
Else
    Range("B" & Application.WorksheetFunction.Match("*Close*", Range("A:A"), 0)).Value = ""
End If
Range("B" & Application.WorksheetFunction.Match("RMA*", Range("A:A"), 0)).Value = rmaNo

'fill in claim table details
tableName = "WarrantyLog"

Set recSet = New ADODB.Recordset
recSet.Open tableName, cN, adOpenKeyset, adLockPessimistic, adCmdTable

recSet.Filter = "Complaint_No = '" & ccNo & "'"

numRecs = recSet.RecordCount 'number of records/table rows

If numRecs = 0 Then
    Exit Function
End If

'copy recordset into column 2 to the left of the table (because index column & complaint column)
Sheet3.Cells(6, Application.WorksheetFunction.Match("Part*", Range("5:5"), 0) - 2).CopyFromRecordset recSet

recSet.Close
Set recSet = Nothing
cN.Close
Set cN = Nothing

Exit Function

errhandler:

MsgBox "Error in ShowAllClaimInfo function"

End Function

Function ClearSheet(rowNum As Integer)

Dim colNum As Integer, i As Integer
Dim wSheet As Worksheet, compCol As Integer, supCol As Integer, rootCol As Integer

On Error GoTo errhandler

Application.ScreenUpdating = False
compCol = 0
supCol = 0
rootCol = 0

Set wSheet = ThisWorkbook.ActiveSheet

wSheet.Activate
wSheet.Unprotect
wSheet.Range("A1:B1").UnMerge
Range("B:B").ClearContents
Range("A1:B1").Merge
Range("D3:D15").ClearContents

If rowNum < 6 Then
    Exit Function
End If

colNum = Application.WorksheetFunction.Match("ID", wSheet.Range("5:5"), 0)

For i = colNum To colNum + 20
    If wSheet.Cells(5, i).Value Like "Complaint*Cat*" Then
        compCol = i
    ElseIf wSheet.Cells(5, i).Value Like "*Supplier*" Then
        supCol = i
    ElseIf wSheet.Cells(5, i).Value Like "Root*Cat*" Then
        rootCol = i
    Else
        Range(Cells(6, i), Cells(75, i)).Select
        Selection.ClearContents
    End If
Next i

If compCol > 0 Then 'delete complaint categories
    Range(Cells(6, compCol), Cells(rowNum, compCol)).ClearContents
End If

If supCol > 0 Then 'delete suppliers
    Range(Cells(6, supCol), Cells(rowNum, supCol)).ClearContents
End If

If rootCol > 0 Then 'delete root cause categories
    Range(Cells(6, rootCol), Cells(rowNum, rootCol)).ClearContents
End If

wSheet.Protect
Range("A1").Select
Exit Function

errhandler:
MsgBox "Error in ClearSheet function"

End Function

Sub ModifyDB()

Dim numRows As Integer, dbPath As String, tableName As String, cN As ADODB.Connection, recSet As ADODB.Recordset
Dim contInfoOnly As Boolean, i As Integer, j As Integer, lastRow As Integer, firstCol As Integer, compCatCol As Integer
Dim supCol As Integer, rcCol As Integer, chkRng As Range, delRow As Boolean, delRows() As Integer, x As Integer
Dim custNum As Integer, custName As String, contName As String, compCat As String, strSupplier As String, strRCCat As String

Call CreateBackupDB

On Error Resume Next
numRows = Application.WorksheetFunction.Match("Ready", Sheet3.Range("A:A"), 0)
On Error GoTo errhandler

If numRows = 0 Then
    MsgBox "Enter more information before proceeding."
    Exit Sub
End If

If Range("F6").Value > 0 Then
    numRows = 6
    If Range("F7").Value > 0 Then
        Range("F6").Select
        Selection.End(xlDown).Select
        numRows = ActiveCell.Row
    End If
Else
    numRows = 0
End If

Sheet3.Activate
lastRow = CheckTable

firstCol = Application.WorksheetFunction.Match("Part*N*", Sheet3.Range("5:5"), 0)
compCatCol = Application.WorksheetFunction.Match("Complain*Cat*", Sheet3.Range("5:5"), 0)
supCol = Application.WorksheetFunction.Match("*Supplier*", Sheet3.Range("5:5"), 0)
rcCol = Application.WorksheetFunction.Match("Root*Cat*", Sheet3.Range("5:5"), 0)

'confirm items to be deleted
x = 0

For i = 6 To Application.WorksheetFunction.Max(numRows, lastRow)
    delRow = True 'delete this record from DB
    For j = firstCol To rcCol + 2
        If j = compCatCol Then
            If Cells(i, j).Value <> Sheet4.Range("M1").Value Then
                delRow = False
            End If
        ElseIf j = supCol Then
            If Cells(i, j).Value <> Sheet4.Range("Q1").Value Then
                delRow = False
            End If
        ElseIf j = rcCol Then
            If Cells(i, j).Value <> Sheet4.Range("O1").Value Then
                delRow = False
            End If
        Else
            If Not (IsEmpty(Cells(i, j))) Then
                delRow = False
            End If
        End If
    Next j
    
    If delRow Then 'add to list of rows to delete records for
        ReDim Preserve delRows(x)
        delRows(x) = i
        x = x + 1
    End If
Next i

If x > 0 Then
    delRow = True 'indicative that deletion needs to occur
    x = x - 1 'x is the index of last item in the list
Else
    delRow = False 'no deletion will be involved
End If

dbPath = Sheet1.Range("A" & Application.WorksheetFunction.Match("Full*D*B*", Sheet1.Range("A:A"), 0) + 1).Value

tableName = "WarrantyLog"

Set cN = New ADODB.Connection
cN.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & dbPath
cN.Open

Set recSet = New ADODB.Recordset
recSet.Open tableName, cN, adOpenKeyset, adLockPessimistic, adCmdTable

Sheet3.Activate

'Update records & add new ones
If numRows = lastRow Then 'no rows added or deleted
    If numRows = 0 Then 'nothing to update on this table
        MsgBox "Only customer's contact info will be updated."
    Else 'update info in records, but don't add or delete any
        For i = 6 To numRows 'update records
            recSet.Filter = "ID = '" & Cells(i, Application.WorksheetFunction.Match("ID", Sheet3.Range("5:5"), 0)).Value & "'"
            If recSet.RecordCount = 1 Then 'overwrite
                recSet.Fields("Complaint") = Range("B" & Application.WorksheetFunction.Match("Complaint*", Sheet3.Range("A:A"), 0)).Value
                recSet.Fields("Part_No") = Cells(i, Application.WorksheetFunction.Match("Part*N*", Sheet3.Range("5:5"), 0)).Value
                recSet.Fields("Serial_No") = Cells(i, Application.WorksheetFunction.Match("Part*SN*", Sheet3.Range("5:5"), 0)).Value
                recSet.Fields("Machine_Model") = Cells(i, Application.WorksheetFunction.Match("Mach*Mod*", Sheet3.Range("5:5"), 0)).Value
                recSet.Fields("Machine_SN") = Cells(i, Application.WorksheetFunction.Match("Mach*SN*", Sheet3.Range("5:5"), 0)).Value
                recSet.Fields("Complaint_Cat") = Cells(i, Application.WorksheetFunction.Match("Complain*Cat*", Sheet3.Range("5:5"), 0)).Value
                recSet.Fields("Complaint") = Cells(i, Application.WorksheetFunction.Match("Complaint", Sheet3.Range("5:5"), 0)).Value
                recSet.Fields("Item_Description") = Cells(i, Application.WorksheetFunction.Match("Item*Desc*", Sheet3.Range("5:5"), 0)).Value
                recSet.Fields("Lot_No") = Cells(i, Application.WorksheetFunction.Match("Lot*", Sheet3.Range("5:5"), 0)).Value
                recSet.Fields("Supplier") = Cells(i, Application.WorksheetFunction.Match("*Supplier*", Sheet3.Range("5:5"), 0)).Value
                recSet.Fields("RootCause_Cat") = Cells(i, Application.WorksheetFunction.Match("Root*Cat*", Sheet3.Range("5:5"), 0)).Value
                recSet.Fields("Root_Cause") = Cells(i, Application.WorksheetFunction.Match("Root*Cause", Sheet3.Range("5:5"), 0)).Value
                recSet.Fields("SCAR") = Cells(i, Application.WorksheetFunction.Match("SCAR", Sheet3.Range("5:5"), 0)).Value
                recSet.Fields("CAPA") = Cells(i, Application.WorksheetFunction.Match("CAPA", Sheet3.Range("5:5"), 0)).Value
                recSet.Update
            Else '0 records found, or more than 1 found
                MsgBox "Error matching table entries to database. Issue with WarrantyLog ID values."
            End If
        Next i
    End If
    
Else 'rows added or rows deleted from the end of the list

    For i = 6 To numRows 'update items currently in DB
        recSet.Filter = "ID = '" & Cells(i, Application.WorksheetFunction.Match("ID", Sheet3.Range("5:5"), 0)).Value & "'"
        If recSet.RecordCount = 1 Then 'overwrite
            recSet.Fields("Complaint") = Range("B" & Application.WorksheetFunction.Match("Complaint*", Sheet3.Range("A:A"), 0)).Value
            recSet.Fields("Part_No") = Cells(i, Application.WorksheetFunction.Match("Part*N*", Sheet3.Range("5:5"), 0)).Value
            recSet.Fields("Serial_No") = Cells(i, Application.WorksheetFunction.Match("Part*SN*", Sheet3.Range("5:5"), 0)).Value
            recSet.Fields("Machine_Model") = Cells(i, Application.WorksheetFunction.Match("Mach*Mod*", Sheet3.Range("5:5"), 0)).Value
            recSet.Fields("Machine_SN") = Cells(i, Application.WorksheetFunction.Match("Mach*SN*", Sheet3.Range("5:5"), 0)).Value
            recSet.Fields("Complaint_Cat") = Cells(i, Application.WorksheetFunction.Match("Complain*Cat*", Sheet3.Range("5:5"), 0)).Value
            recSet.Fields("Complaint") = Cells(i, Application.WorksheetFunction.Match("Complaint", Sheet3.Range("5:5"), 0)).Value
            recSet.Fields("Item_Description") = Cells(i, Application.WorksheetFunction.Match("Item*Desc*", Sheet3.Range("5:5"), 0)).Value
            recSet.Fields("Lot_No") = Cells(i, Application.WorksheetFunction.Match("Lot*", Sheet3.Range("5:5"), 0)).Value
            recSet.Fields("Supplier") = Cells(i, Application.WorksheetFunction.Match("*Supplier*", Sheet3.Range("5:5"), 0)).Value
            recSet.Fields("RootCause_Cat") = Cells(i, Application.WorksheetFunction.Match("Root*Cat*", Sheet3.Range("5:5"), 0)).Value
            recSet.Fields("Root_Cause") = Cells(i, Application.WorksheetFunction.Match("Root*Cause", Sheet3.Range("5:5"), 0)).Value
            recSet.Fields("SCAR") = Cells(i, Application.WorksheetFunction.Match("SCAR", Sheet3.Range("5:5"), 0)).Value
            recSet.Fields("CAPA") = Cells(i, Application.WorksheetFunction.Match("CAPA", Sheet3.Range("5:5"), 0)).Value
            recSet.Update
        Else '0 records found, or more than 1 found
            MsgBox "Error matching table entries to database. Issue with WarrantyLog ID values."
        End If
    Next i
    
    If lastRow > numRows Then 'new records need to be added
        For i = numRows + 1 To lastRow
                compCat = Cells(i, Application.WorksheetFunction.Match("Complain*Cat*", Sheet3.Range("5:5"), 0)).Value
                strSupplier = Cells(i, Application.WorksheetFunction.Match("*Supplier*", Sheet3.Range("5:5"), 0)).Value
                If strSupplier = Sheet4.Range("Q1").Value Then
                    strSupplier = ""
                End If
                strRCCat = Cells(i, Application.WorksheetFunction.Match("Root*Cat*", Sheet3.Range("5:5"), 0)).Value
                If strRCCat = Sheet4.Range("O1").Value Then
                    strRCCat = ""
                End If
                
                If compCat <> Sheet4.Range("M1").Value Then
                    recSet.AddNew
                    recSet.Fields("Complaint_No") = Range("B" & Application.WorksheetFunction.Match("Complaint*", Sheet3.Range("A:A"), 0)).Value
                    recSet.Fields("Part_No") = Cells(i, Application.WorksheetFunction.Match("Part*N*", Sheet3.Range("5:5"), 0)).Value
                    recSet.Fields("Serial_No") = Cells(i, Application.WorksheetFunction.Match("Part*SN*", Sheet3.Range("5:5"), 0)).Value
                    recSet.Fields("Machine_Model") = Cells(i, Application.WorksheetFunction.Match("Mach*Mod*", Sheet3.Range("5:5"), 0)).Value
                    recSet.Fields("Machine_SN") = Cells(i, Application.WorksheetFunction.Match("Mach*SN*", Sheet3.Range("5:5"), 0)).Value
                    recSet.Fields("Complaint_Cat") = Cells(i, Application.WorksheetFunction.Match("Complain*Cat*", Sheet3.Range("5:5"), 0)).Value
                    recSet.Fields("Complaint") = Cells(i, Application.WorksheetFunction.Match("Complaint", Sheet3.Range("5:5"), 0)).Value
                    recSet.Fields("Item_Description") = Cells(i, Application.WorksheetFunction.Match("Item*Desc*", Sheet3.Range("5:5"), 0)).Value
                    recSet.Fields("Lot_No") = Cells(i, Application.WorksheetFunction.Match("Lot*", Sheet3.Range("5:5"), 0)).Value
                    recSet.Fields("Supplier") = strSupplier 'might need to be changed to null
                    recSet.Fields("RootCause_Cat") = strRCCat 'might need to be changed to null
                    recSet.Fields("Root_Cause") = Cells(i, Application.WorksheetFunction.Match("Root*Cause", Sheet3.Range("5:5"), 0)).Value
                    recSet.Fields("SCAR") = Cells(i, Application.WorksheetFunction.Match("SCAR", Sheet3.Range("5:5"), 0)).Value
                    recSet.Fields("CAPA") = Cells(i, Application.WorksheetFunction.Match("CAPA", Sheet3.Range("5:5"), 0)).Value
                    recSet.Update
                Else
                    MsgBox "Row " & i & " of the table has not been added to the database because it doesn't include a complaint category."
                End If
        Next i
        
    End If
End If

'delete any records that were left blank in the middle of the list
If delRow = True Then 'at least one record needs to be deleted
    For Each rowID In delRows
        recSet.Filter = "ID = '" & Cells(rowID, Application.WorksheetFunction.Match("ID", Sheet3.Range("5:5"), 0)).Value & "'"
        If recSet.RecordCount = 1 Then
            recSet.Delete
            recSet.Update
        Else
            MsgBox "Error locating (in the database) & deleting the record in row " & rowID
        End If
    Next rowID
End If

recSet.Close

'update customer info
tableName = "Customers"
Set recSet = New ADODB.Recordset
recSet.Open tableName, cN, adOpenKeyset, adLockPessimistic, adCmdTable

custName = Sheet3.Range("B" & Application.WorksheetFunction.Match("Customer*", Sheet3.Range("A:A"), 0)).Value
contName = Sheet3.Range("B" & Application.WorksheetFunction.Match("Contact*", Sheet3.Range("A:A"), 0)).Value
recSet.Filter = "Customer_Name = '" & custName & "'"

If recSet.RecordCount = 1 Then 'update customer
    recSet.Fields("Customer_Name") = Sheet3.Range("B" & Application.WorksheetFunction.Match("Customer*", Sheet3.Range("A:A"), 0)).Value
    custNum = recSet.Fields("ID")
    recSet.Update

ElseIf recSet.RecordCount = 0 Then 'add customer & change claim table to reflect new customer
    recSet.AddNew 'add customer
    recSet.Fields("Customer_Name") = Sheet3.Range("B" & Application.WorksheetFunction.Match("Customer*", Sheet3.Range("A:A"), 0)).Value
    custNum = recSet.Fields("ID")
    recSet.Update
    
Else
    MsgBox "Error locating the customer/contact person pair. Contact info has not been updated."
End If

recSet.Close

'update contact info
tableName = "Contacts"
Set recSet = New ADODB.Recordset
recSet.Open tableName, cN, adOpenKeyset, adLockPessimistic, adCmdTable

recSet.Filter = "Customer = '" & custNum & "'AND Contact = '" & contName & "'"

If recSet.RecordCount = 1 Then 'update customer
    recSet.Fields("Contact") = Sheet3.Range("B" & Application.WorksheetFunction.Match("Contact*", Sheet3.Range("A:A"), 0)).Value
    recSet.Fields("Customer") = custNum
    recSet.Fields("Address") = Sheet3.Range("B" & Application.WorksheetFunction.Match("Address*", Sheet3.Range("A:A"), 0)).Value
    recSet.Fields("City") = Sheet3.Range("B" & Application.WorksheetFunction.Match("City*", Sheet3.Range("A:A"), 0)).Value
    recSet.Fields("State") = Sheet3.Range("B" & Application.WorksheetFunction.Match("State*", Sheet3.Range("A:A"), 0)).Value
    recSet.Fields("ZIP") = Sheet3.Range("B" & Application.WorksheetFunction.Match("Zip*", Sheet3.Range("A:A"), 0)).Value
    recSet.Fields("Country") = Sheet3.Range("B" & Application.WorksheetFunction.Match("Country*", Sheet3.Range("A:A"), 0)).Value
    recSet.Fields("Phone") = Sheet3.Range("B" & Application.WorksheetFunction.Match("Phone*", Sheet3.Range("A:A"), 0)).Value
    recSet.Fields("Email") = Sheet3.Range("B" & Application.WorksheetFunction.Match("Email*", Sheet3.Range("A:A"), 0)).Value
    custNum = recSet.Fields("ID")
    recSet.Update

ElseIf recSet.RecordCount = 0 Then 'add customer & change claim table to reflect new customer
    recSet.AddNew 'add customer
    recSet.Fields("Contact") = Sheet3.Range("B" & Application.WorksheetFunction.Match("Contact*", Sheet3.Range("A:A"), 0)).Value
    recSet.Fields("Customer") = custNum
    recSet.Fields("Address") = Sheet3.Range("B" & Application.WorksheetFunction.Match("Address*", Sheet3.Range("A:A"), 0)).Value
    recSet.Fields("City") = Sheet3.Range("B" & Application.WorksheetFunction.Match("City*", Sheet3.Range("A:A"), 0)).Value
    recSet.Fields("State") = Sheet3.Range("B" & Application.WorksheetFunction.Match("State*", Sheet3.Range("A:A"), 0)).Value
    recSet.Fields("ZIP") = Sheet3.Range("B" & Application.WorksheetFunction.Match("Zip*", Sheet3.Range("A:A"), 0)).Value
    recSet.Fields("Country") = Sheet3.Range("B" & Application.WorksheetFunction.Match("Country*", Sheet3.Range("A:A"), 0)).Value
    recSet.Fields("Phone") = Sheet3.Range("B" & Application.WorksheetFunction.Match("Phone*", Sheet3.Range("A:A"), 0)).Value
    recSet.Fields("Email") = Sheet3.Range("B" & Application.WorksheetFunction.Match("Email*", Sheet3.Range("A:A"), 0)).Value
    custNum = recSet.Fields("ID")
    recSet.Update
    
Else
    MsgBox "Error locating the customer/contact person pair. Contact info has not been updated."
End If


'update claim info
tableName = "ClaimInfo"
Set recSet = New ADODB.Recordset
recSet.Open tableName, cN, adOpenKeyset, adLockPessimistic, adCmdTable

recSet.Filter = "Complaint_No = '" & Sheet3.Range("B" & Application.WorksheetFunction.Match("Complaint*", Sheet3.Range("A:A"), 0)).Value & "'"

If recSet.RecordCount = 1 Then
    recSet.Fields("Initiated_By") = Sheet3.Range("B" & Application.WorksheetFunction.Match("Quality*", Sheet3.Range("A:A"), 0)).Value
    recSet.Fields("CustomerContact") = custNum
    recSet.Fields("Date_Opened") = Sheet3.Range("B" & Application.WorksheetFunction.Match("Date*Open*", Sheet3.Range("A:A"), 0)).Value
    recSet.Fields("RMA_No") = Sheet3.Range("B" & Application.WorksheetFunction.Match("RMA*", Sheet3.Range("A:A"), 0)).Value
    recSet.Fields("Date_Closed") = Sheet3.Range("B" & Application.WorksheetFunction.Match("Date*Close*", Sheet3.Range("A:A"), 0)).Value
    recSet.Update
Else
    MsgBox "Error locating claim number in the database. Complaint info has not been updated."
End If

MsgBox "Database updated successfully."

Call ClearSheet(Application.WorksheetFunction.Max(numRows, lastRow))
Call BackToMain

Exit Sub

errhandler:
MsgBox "Error in ModifyDB sub"

End Sub
