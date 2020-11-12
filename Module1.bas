Attribute VB_Name = "Module1"
Sub NewEntry()

'Sheet1 -> Sheet2

If InStr(UCase(ThisWorkbook.Path), "ENGLANDT") = 0 Then
    MsgBox "Please use the correct workbook, not a copy."
    Exit Sub
End If

Dim rowNum As Integer

Application.ScreenUpdating = False
rowNum = Application.WorksheetFunction.Match("D*B*Loc*", Range("A:A"), 0) + 1

Call CheckDB 'undoes any manual highlighting

If Range("A" & rowNum).Interior.ColorIndex <> -4142 Then
    If Range("A" & rowNum).Interior.ColorIndex = 3 Then
        Exit Sub 'error message thrown by checkDB above
    ElseIf Range("A" & rowNum).Interior.ColorIndex = 6 Then
        Exit Sub 'error message thrown by checkDB above
    Else
        MsgBox "An issue exists with the database and/or its directory. Repair this issue before proceeding."
    End If
    
    Exit Sub
End If

Sheet2.Visible = True
Sheet2.Activate
Range("A1").Select
Sheet1.Visible = False

Application.ScreenUpdating = True

End Sub


Sub UpdateEntry()

'input box -> Sheet3

If InStr(UCase(ThisWorkbook.Path), "ENGLANDT") = 0 Then
    MsgBox "Please use the correct workbook, not a copy."
    'Exit Sub
End If

Dim rowNum As Integer, claimNo As String, claimPresent As Boolean, preExClaim As String

Application.ScreenUpdating = False

On Error GoTo errhandler

rowNum = Application.WorksheetFunction.Match("D*B*Loc*", Range("A:A"), 0) + 1

Call CheckDB 'undoes any manual highlighting

If Range("A" & rowNum).Interior.ColorIndex <> -4142 Then
    If Range("A" & rowNum).Interior.ColorIndex = 3 Then
        Exit Sub 'error message thrown by checkDB above
    ElseIf Range("A" & rowNum).Interior.ColorIndex = 6 Then
        Exit Sub 'error message thrown by checkDB above
    Else
        MsgBox "An issue exists with the database and/or its directory. Repair this issue before proceeding."
    End If
    
    Exit Sub
End If

preExClaim = Sheet3.Range("B" & Application.WorksheetFunction.Match("Complaint*", Sheet3.Range("A:A"), 0)).Value

If preExClaim > "" Then
    'info was already entered on the 'update DB' sheet
    claimNo = MsgBox("Do you want to continue modifying claim " & preExClaim & "?", vbYesNo)
    If claimNo = vbYes Then 'no checking, just change sheets
        Sheet3.Visible = True
        Sheet3.Activate
        Sheet1.Visible = False
        Range("A1").Select
        Exit Sub
    End If
End If

claimNo = InputBox("Enter the customer complaint number whose record you want to modify (format 'CCXX-XXX')", "CC Number")

If claimNo = "" Then
    Exit Sub
Else
    claimNo = UCase(claimNo)
End If

'check for cc records
claimPresent = CheckForClaim(claimNo)

If claimPresent Then
    Sheet3.Visible = True
    Sheet3.Activate
    Range("A1").Select
    Sheet1.Visible = False
Else
    MsgBox "That claim number was not found in the database."
    Exit Sub
End If

If preExClaim > "" Then 'get rid of old info
    Call ClearSheet(70)
    Range("E105:S175").ClearContents 'clear old ref info
End If

'populate info
Call ShowAllClaimInfo(claimNo)

Application.ScreenUpdating = False
Range("E6:S75").Copy
Range("E106:S175").PasteSpecial xlPasteValues
Application.CutCopyMode = False
Range("A1").Select

Application.ScreenUpdating = True

Exit Sub

errhandler:
MsgBox "Error in UpdateEntry sub"
Application.ScreenUpdating = True

End Sub


Sub BackToMain()

'Activesheet -> Sheet1

If ActiveSheet.Name = Sheet1.Name Then
    Exit Sub
End If

Dim wSheet As Worksheet
Application.ScreenUpdating = False
Set wSheet = ActiveSheet
Sheet1.Visible = True
Sheet1.Activate
Range("A1").Select
wSheet.Visible = False
Application.ScreenUpdating = True

End Sub


Sub AddEntryToDB()

Dim i As Integer, j As Integer, k As Integer, deleteEntry As Boolean, custName As String, rowNum As Integer, custRows() As Integer
Dim cContact As String, cAddress As String, cCity As String, cState As String, cCountry As String, cZIP As String, x As Boolean

Application.ScreenUpdating = False

Call CreateBackupDB

ThisWorkbook.Activate
Sheet2.Activate

On Error Resume Next
i = Application.WorksheetFunction.Match("Ready", Range("A:A"), 0)
On Error GoTo errhandler

If i = 0 Then
    MsgBox "Enter more information before proceeding."
    Exit Sub
End If

x = CheckDB 'checkdb is false if db is renamed/moved
If Not (x) Then
    MsgBox "Issue with the database file and/or location. Ensure that " & _
            "the file has not been moved or renamed."
    Exit Sub
End If

If i = 0 Then 'status is not "Ready"
    Set cityRng = Range("A" & Application.WorksheetFunction.Match("City*", Range("A:A"), 0)).Offset(0, 1)
    Set stateRng = Range("A" & Application.WorksheetFunction.Match("State*", Range("A:A"), 0)).Offset(0, 1)
    Set zipRng = Range("A" & Application.WorksheetFunction.Match("Zip*", Range("A:A"), 0)).Offset(0, 1)
    
    If cityRng.Value = 0 And cityRng.Offset(0, 2).Value = False Then
        msgAns = MsgBox("No city has been entered. Do you want to submit this entry without a city?", vbYesNo, "City Error")
        If msgAns = vbYes Then
            cityRng.Offset(0, 2).Value = "TRUE"
            MsgBox "No city will be required. Please retry submitting the new complaint."
        End If
        Exit Sub
    ElseIf stateRng.Value = 0 And stateRng.Offset(0, 2).Value = False Then
        msgAns = MsgBox("No address state has been entered. Do you want to submit this entry without a state?", vbYesNo, "State Error")
        If msgAns = vbYes Then
            stateRng.Offset(0, 2).Value = "TRUE"
            MsgBox "No state will be required. Please retry submitting the new complaint."
        End If
        Exit Sub
    ElseIf zipRng.Value = 0 And zipRng.Offset(0, 2).Value = False Then
        msgAns = MsgBox("No zip code has been entered. Do you want to submit this entry without a zip code?", vbYesNo, "Zip Code Error")
        If msgAns = vbYes Then
            zipRng.Offset(0, 2).Value = "TRUE"
            MsgBox "No zip code will be required. Please retry submitting the new complaint."
        End If
        Exit Sub
    Else
        MsgBox ("Additional info is required before the entry can be added to the database.")
        Exit Sub
    End If
Else 'status is "Ready"
    msgAns = MsgBox("Ensure all relevant info has been added to the table on the right!" & vbCrLf & _
               vbCrLf & "Do you want to proceed?", vbYesNo)
    If msgAns = vbNo Then
        Exit Sub
    End If
End If

Sheet2.Activate
i = CheckTable

If i = 0 Then
    Exit Sub
End If

'update customer list with entered info
x = UpdateCustomerInfo
If Not x Then
    Exit Sub
End If

'add complaint to complaint table
x = AddToComplaintTable
If Not x Then
    Exit Sub
End If

'add items to warranty table
If i > 5 Then
    x = AddToWarrantyTable(i)
    If Not x Then
        Exit Sub
    End If
End If

MsgBox "Database updated successfully"

Sheet2.Activate
Call ClearSheet(i)
Call BackToMain

Application.ScreenUpdating = True

Exit Sub
errhandler:
    MsgBox "Error in AddEntryToDB sub"
    Application.ScreenUpdating = True
    Call ErrorRep("AddEntryToDB", "Sub", "N/A", Err.Number, Err.Description, "")
End Sub

Function CheckTable() As Integer

Dim i As Integer, j As Integer, k As Integer, cityRng As Range, stateRng As Range, chkRng As Range, wSheet As Worksheet
Dim col1 As Integer, catCol As Integer, supCol As Integer, rcCol As Integer, x As Integer, errorMsg As Boolean


On Error GoTo errhandler

Application.ScreenUpdating = False

CheckTable = 0

Set wSheet = ThisWorkbook.ActiveSheet

'check for rows with info entered but without complaint categories
catCol = Application.WorksheetFunction.Match("Complaint*Cat*", wSheet.Range("5:5"), 0) 'column with comp cat

supCol = Application.WorksheetFunction.Match("*Supplier*", wSheet.Range("5:5"), 0) 'column with supplier
col1 = Application.WorksheetFunction.Match("Part*Num*", wSheet.Range("5:5"), 0) 'first column of table (LH)
On Error Resume Next
rcCol = 0
rcCol = Application.WorksheetFunction.Match("*Root*Cat*", wSheet.Range("5:5"), 0) 'column with supplier
On Error GoTo errhandler

For i = 75 To 4 Step -1 'find last complaint cat row (not counting "Complaint Categories"
    If wSheet.Cells(i, catCol).Value <> 0 And wSheet.Cells(i, catCol).Value <> Sheet4.Range("P1").Value Then
        Exit For
    End If
Next i

i = i + 1 'row i is first row after last category entry

Set chkRng = Range(Cells(i, col1), Cells(75, catCol - 1))

errorMsg = False 'False = range is empty
For Each cell In chkRng 'cells below i in table to left of category column
    errorMsg = Not (IsEmpty(cell))
    If errorMsg = True Then 'cell isn't empty
        Exit For
    End If
Next cell

If errorMsg = False Then
    Set chkRng = Range(Cells(i, catCol + 1), Cells(75, supCol - 1))
    For Each cell In chkRng 'cells below i in table between category & supplier columns
        errorMsg = Not (IsEmpty(cell))
        If errorMsg Then 'cell isn't empty
            Exit For
        End If
    Next
End If

If errorMsg = False And rcCol > 0 Then
    Set chkRng = Range(Cells(i, rcCol + 1), Cells(75, rcCol + 3))
    For Each cell In chkRng 'cells below i in table to right of root cause cat column
        errorMsg = Not (IsEmpty(cell))
        If errorMsg Then 'cell isn't empty
            Exit For
        End If
    Next
End If

If errorMsg = False Then
    Set chkRng = Range(Cells(i, supCol), Cells(75, supCol)) 'cells in supplier column
    For Each cell In chkRng 'cells in supplier column
        If cell.Value <> Sheet4.Range("T1").Value Then
            errorMsg = True
            Exit For
        End If
    Next
End If

If errorMsg = False And rcCol > 0 Then
    Set chkRng = Range(Cells(i, rcCol), Cells(75, rcCol)) 'cells in rootcause category column
    For Each cell In chkRng 'cells in supplier column
        If cell.Value <> Sheet4.Range("R1").Value Then
            errorMsg = True
            Exit For
        End If
    Next
End If

i = i - 1 'i=last row of interest

If errorMsg Then 'data in a row after last complaint category
    MsgBox "A Complaint Category is required for each row of the table that has data for the database. " & _
            "Please enter the missing Complaint Categories and retry."
    Exit Function
End If

If i = 5 Then 'no complaint categories or data
    If rcCol = 0 Then
        CheckTable = 4
        MsgBox "This customer complaint will be added to the database, but no parts will be attributed to it."
    End If
    
    Exit Function
    
Else 'at least one row of the table is populated

    'convert all "Item Description" cells to capital letters
    j = Application.WorksheetFunction.Match("*Description*", wSheet.Range("5:5"), 0)
    For k = 6 To i
        If Cells(k, j).Value <> 0 Then
            Cells(k, j).Value = UCase(Cells(k, j).Value)
        End If
    Next k

    'make sure no cat's were skipped
    For k = 6 To i
        If Cells(k, catCol).Value = 0 Or Cells(k, catCol).Value = Sheet4.Range("M1").Value Then
            If rcCol = 0 Then 'rccol>0 allows category to be skipped (ModifyDB)
                errorMsg = True
                Exit For
            End If
        End If
    Next k
End If

If errorMsg Then
    MsgBox "The last row of the table with a Complaint Category is row " & i & ". Please " & _
            "ensure that all rows before that one have Complaint Categories as well, and then retry."
    Exit Function
End If

CheckTable = i

Exit Function

errhandler:
    MsgBox "Error in CheckTable function"
    Call ErrorRep("CheckTable", "Function", CheckTable, Err.Number, Err.Description, "")
End Function

Function AlphaSort(newText As String, exText As String) As Boolean
                    'too               'tool
Dim minLen As Integer, i As Integer, textNew As String, textOld As String

On Error GoTo errhandler

Application.ScreenUpdating = False

minLen = WorksheetFunction.Min(Len(newText), Len(exText))
textNew = newText 'these can be mangled without affecting the actual text
textOld = exText 'these can be mangled without affecting the actual text
AlphaSort = False
i = 0

While i = 0
    If Left(textNew, 1) < Left(textOld, 1) Then
        AlphaSort = True 'new text comes first
        i = 1
    ElseIf Left(textNew, 1) > Left(textOld, 1) Then
        AlphaSort = False 'old text comes first
        i = 1
    Else
        textNew = Right(textNew, Len(textNew) - 1) 'next letter
        textOld = Right(textOld, Len(textOld) - 1) 'next letter
    End If
    
    If Len(textNew) = 0 Or Len(textOld = 0) Then
        i = 5 'one word is part of the other word
    End If
Loop

If x = 5 Then 'shorter word comes first
    If Len(newText) < Len(exText) Then
        AlphaSort = True
    End If
End If

Exit Function
errhandler:
    MsgBox "Error in AlphaSort function"
    Call ErrorRep("AlphaSort", "Function", AlphaSort, Err.Number, Err.Description, newText & "   " & exText)
End Function

Sub NewComplaintCat()
'input boxes -> Add to Complaint category list

Dim newCompCat As String

Application.ScreenUpdating = False

newCompCat = InputBox("Enter a name for the new complaint category", "New Complaint Category")

If newCompCat = "" Then
    Exit Sub
End If

Call AddToList("Complaint", newCompCat)

End Sub

Sub NewRootCauseCat()
'input boxes -> Add to Root Cause category list

Dim newRCCat As String

Application.ScreenUpdating = False

newRCCat = InputBox("Enter a name for the new Root Cause category", "New Root Cause Category")

If newRCCat = "" Then
    Exit Sub
End If

Call AddToList("Cause", newRCCat)

End Sub

Sub NewSupplier()
'input boxes -> Add to Supplier list

Dim NewSupplier As String

Application.ScreenUpdating = False

NewSupplier = InputBox("Enter the supplier name to be added", "Add Supplier To List")

If NewSupplier = "" Then
    Exit Sub
End If

Call AddToList("Supplier", NewSupplier)

End Sub

Sub AddToList(listKeyword As String, newText As String)

Dim wSheet As Worksheet, colNum As Integer, rowNum As Integer, i As Integer, k As Integer

Application.ScreenUpdating = False

On Error GoTo errhandler

Set wSheet = ActiveSheet
Sheet4.Activate

colNum = Application.WorksheetFunction.Match("*" & listKeyword & "*", Sheet4.Range("1:1"), 0) 'Identify column of list

i = 0
On Error Resume Next
i = Application.WorksheetFunction.Match(newText, Range(colNum & ":" & colNum), 0) 'check that it isn't already there
On Error GoTo errhandler

If i > 0 Then
    MsgBox "It seems this text is already a part of the list."
    Exit Sub
End If

Cells(2, colNum).Select 'Identify row where new item belongs
Selection.End(xlDown).Select
i = ActiveCell.Row
For k = 2 To i
    If AlphaSort(newText, Cells(k, colNum).Value) Then 'k is row to insert
        Cells(k, colNum).Select
        Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove 'insert row
        rowNum = k
        k = 500 'indicative that placement has occurred (not in last row)
        Exit For
    End If
Next k

If k < 500 Then 'new item is appended at the end of the list
    rowNum = i + 1
End If

Cells(rowNum, colNum).Value = newText 'Add new item to list

Cells(rowNum + 1, colNum).Copy
Cells(rowNum, colNum).PasteSpecial xlFormats
Application.CutCopyMode = False
Cells(1, 1).Select

wSheet.Activate
MsgBox newText & " successfully added! " & newText & " can now be chosen from the dropdown."

Exit Sub
errhandler:
    MsgBox "Errror in AddToList function"
    Call ErrorRep("AddToList", "Sub", "N/A", Err.Number, Err.Description, "")
End Sub
