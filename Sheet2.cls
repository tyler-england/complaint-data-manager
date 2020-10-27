VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Private Sub Worksheet_Activate()
'Adds complaint number and user's name

Dim ccNoRng As Range, userName As Range, custName As Range, dateRange As Range
Dim ccNoOld As String, ccNo As String, i As Integer
Dim uName As String, fName As String, lName As String

Set ccNoRng = Range("B" & Application.WorksheetFunction.Match("Complaint*", Range("A:A"), 0))
Set userName = ccNoRng.Offset(1, 0)
Set custName = userName.Offset(1, 0)
Set dateRange = Range("B" & Application.WorksheetFunction.Match("*Date*", Range("A:A"), 0))

'Make sure DB is linked
ThisWorkbook.RefreshAll

'Add complaint number
If ccNoRng.Value = 0 Then
    'check Access database?
    'Excel CC log prob going away
    'check folders?
End If
''''

'Add user's name
If userName.Value = 0 Then
    uName = Application.userName
    lName = Left(uName, InStr(uName, ",") - 1)
    uName = Right(uName, Len(uName) - InStr(uName, " "))
    fName = Left(uName, InStr(uName, " "))
    userName.Value = fName & lName
End If
''''

'Add customer dropdown value
If custName.Value = 0 Then
    custName.Value = "Customer_Name"
End If
''''

'Add date
If dateRange.Value = 0 Then
    dateRange.Value = Date
End If
''''

End Sub


Private Sub Worksheet_Change(ByVal Target As Range)
'Updates a lot of things all the time

Dim compRange As Range, custRange As Range, custContRng As Range, stateRange As Range, cntryRange As Range
Dim lookupRow As Integer, i As Integer, j As Integer, custRangePop As Range, custContRngPop As Range
Dim cName As String, cContact As String, cAddress As String, cCity As String, capVal As String, custID As Integer
Dim cState As String, cZIP As String, cCountry As String, stateInput As String, qtyInList As Integer

j = Application.WorksheetFunction.Match("Comp*Cat*", Range("5:5"), 0)
For i = 5 To 100
    If Cells(i, j).Value = 0 Then
        Cells(i, j).Value = "Complaint Categories" 'keeps categories from being blank
    End If
Next i

j = Application.WorksheetFunction.Match("*Supplier*", Range("5:5"), 0)
For i = 5 To 100
    If Cells(i, j).Value = 0 Then
        Cells(i, j).Value = "Suppliers" 'keeps supplier from being blank
    End If
Next i

Set compRange = Range("B" & Application.WorksheetFunction.Match("Complaint*", Range("A:A"), 0))
Set custRange = Range("B" & Application.WorksheetFunction.Match("Customer*", Range("A:A"), 0))
Set custContRng = Range("B" & Application.WorksheetFunction.Match("Contact*", Range("A:A"), 0))
Set stateRange = Range("B" & Application.WorksheetFunction.Match("State*", Range("A:A"), 0))
Set cntryRange = Range("B" & Application.WorksheetFunction.Match("Country*", Range("A:A"), 0))

If Not Application.Intersect(compRange, Range(Target.Address)) Is Nothing Then
'capitalize complaint number
    If compRange.Value = UCase(compRange.Value) Then
        Exit Sub
    End If
    If compRange > 0 Then
        compRange.Value = UCase(compRange.Value)
    End If
    
ElseIf Not Application.Intersect(custRange, Range(Target.Address)) Is Nothing Then
'Add correct data validation dropdown for Contact

    Sheet2.Unprotect

    If custRange.Value = 0 Then 'no customer listed
        custRange.Value = "Customer_Name"
        If custContRng.Value = "Contact Name" Then
            custContRng.ClearContents
        End If
    ElseIf custRange.Value = "Customer_Name" Then 'no customer listed
        custRange.Font.Italic = True
        If custContRng.Value = "Contact Name" Then
            custContRng.ClearContents
        End If
    Else 'customer listed
        custRange.Font.Italic = False
        
        'correct capitalization if necessary
        i = 0
        For lookupRow = 2 To 500
            If UCase(Sheet4.Range("B" & lookupRow).Value) = UCase(custRange.Value) Then
                If Sheet4.Range("B" & lookupRow).Value = custRange.Value Then
                    Exit For 'capitalization doesn't need fixing
                Else
                    custRange.Value = Sheet4.Range("B" & lookupRow).Value 'replace customer name
                    Exit Sub
                End If
            ElseIf Sheet4.Range("B" & lookupRow).Value = 0 Then
                i = i + 1
                If i > 1 Then
                    Exit For 'new customer name
                End If
            End If
        Next lookupRow
        
        If i > 0 Then 'new customer name
            capVal = CheckCapitalization(custRange.Value, "Customer")
            If custRange.Value <> capVal Then 'change customer name
                custRange.Value = capVal
                Exit Sub
            End If
        End If
        
        custContRng.Validation.Delete
        'check contact info
        If custContRng.Value = 0 Or custContRng.Value = "Contact Name" Or Len(custContRng.Value) < 3 Then 'contact not listed
                'see how many entries are listed for this company name
            If i > 0 Then 'new customer
                Exit Sub
            End If
        
            custID = Application.WorksheetFunction.Index(Sheet4.Range("A:B"), _
                     Application.WorksheetFunction.Match(custRange.Value, Sheet4.Range("B:B"), 0), 1)
        
            qtyInList = Application.WorksheetFunction.CountIf(Sheet4.Range("D:D"), custID)
                'COUNTIF(D:D,INDEX(A:B,MATCH("Merck",B:B,0),1)) would give number of Merck entries
            custContRng.Value = "Contact Name"
            
            If qtyInList = 0 Then
                custContRng.Value = "Contact Name"
            ElseIf qtyInList = 1 Then 'fill in the corresponding contact name
                For i = 1 To 100
                    If Sheet4.Range("D" & i).Value = custID Then
                        lookupRow = i
                        Exit For
                    End If
                Next i
                custContRng.Value = Sheet4.Range("C" & lookupRow).Value
            Else 'set contact name data validation
                Call SetValidation(custID, custContRng)
            End If
        Else 'customer & contact both listed
            If i = 0 Then 'not new customer
                custID = Application.WorksheetFunction.Index(Sheet4.Range("A:B"), _
                     Application.WorksheetFunction.Match(custRange.Value, Sheet4.Range("B:B"), 0), 1)
                 
                Call PopulateCustomerDeets(custID, custContRng)
            End If
        End If
    End If
    
    If custRange.Value <> "Customer_Name" Then
        custRange.Offset(1, 0).Select
    Else
        custRange.Select
    End If
    
    Sheet2.Protect

ElseIf Not Application.Intersect(custContRng, Range(Target.Address)) Is Nothing Then
'customer info populated from customer name & contact

    Sheet2.Unprotect
    
    If custContRng.Value = 0 Then 'no contact listed
        If custRange.Value > 0 And custRange.Value <> "Customer_Name" Then
            custContRng.Value = "Contact Name"
        End If
    ElseIf custContRng.Value = "Contact Name" Then 'no contact listed
        If custRange.Value = "Customer_Name" Then
            custContRng.ClearContents
        Else
            custContRng.Font.Italic = True
        End If
    Else 'contact listed
        custContRng.Font.Italic = False
        
        'correct capitalization if necessary
        i = 0
        For lookupRow = 2 To 500
            If UCase(Sheet4.Range("C" & lookupRow).Value) = UCase(custContRng.Value) Then
                If Sheet4.Range("C" & lookupRow).Value = custContRng.Value Then
                    Exit For 'capitalization doesn't need fixing
                Else
                    custContRng.Value = Sheet4.Range("C" & lookupRow).Value 'replace customer name
                    Exit Sub
                End If
            ElseIf Sheet4.Range("C" & lookupRow).Value = 0 Then
                i = i + 1
                If i > 1 Then
                    Exit For 'new customer name
                End If
            End If
        Next lookupRow
        
        If i > 0 Then 'new contact name
            capVal = CheckCapitalization(custContRng.Value, "Contact")
            If custContRng.Value <> capVal Then 'change contact name
                custContRng.Value = capVal
                Exit Sub
            End If
        End If
        
        'check customer info
        
        If custRange.Value = 0 Or custRange.Value = "Customer_Name" Or Len(custRange.Value) < 3 Then 'customer not listed
            'see how many entries are listed for this contact name
            
            If i > 0 Then 'new customer
                Exit Sub
            End If
            
            qtyInList = Application.WorksheetFunction.CountIf(Sheet4.Range("C:C"), custContRng.Value)

            If qtyInList = 1 Then 'fill in the corresponding customer name
                lookupRow = Application.WorksheetFunction.Match(custContRng.Value, Sheet4.Range("C:C"), 0)
                custID = Sheet4.Range("D" & lookupRow).Value
            
                custRange.Value = Application.WorksheetFunction.Index(Sheet4.Range("A:B"), _
                                  Application.WorksheetFunction.Match(custID, Sheet4.Range("A:A"), 0), 2)
            End If
        Else 'customer is listed
            On Error Resume Next
            custID = Application.WorksheetFunction.Index(Sheet4.Range("A:B"), _
                     Application.WorksheetFunction.Match(custRange.Value, Sheet4.Range("B:B"), 0), 1)
            If custID = 0 Then
                Exit Sub
            End If
            
            Call PopulateCustomerDeets(custID, custContRng)
        End If
    End If
    
    Sheet2.Protect
    custContRng.Offset(1, 0).Select

ElseIf Not Application.Intersect(stateRange, Range(Target.Address)) Is Nothing Then
'check if state was spelled out

    If UCase(cntryRange.Value) = "USA" Or UCase(cntryRange.Value) = "US" Or _
    UCase(cntryRange.Value) = "U.S.A." Or UCase(cntryRange.Value) = "U.S." Or _
    UCase(cntryRange.Value) = "AMERICA" Then 'Country is USA
        If cntryRange.Value <> "USA" Then
            cntryRange.Value = "USA" 'replace whatever it is with "USA"
        End If
        stateInput = UCase(stateRange.Value) 'redefine stateinput
        If Len(stateInput) < 3 Then 'an abbreviation
            i = Application.WorksheetFunction.Match("State*Abbr*", Sheet4.Range("1:1"), 0)
            For lookupRow = 2 To 75
                If Sheet4.Cells(lookupRow, i).Value = stateInput Then
                    Exit For 'state was found
                ElseIf Sheet4.Cells(lookupRow, i).Value = 0 Then
                    lookupRow = 0 'nothing in the list
                    Exit For
                End If
            Next lookupRow
            
            If lookupRow > 0 Then 'state needs to be replaced
                cState = Sheet4.Cells(lookupRow, i + 1).Value
                If Len(cState) < 3 Then
                    Exit Sub 'to repeat infinite loop
                End If
                stateRange.Value = cState
            End If
            
        End If
    End If

ElseIf Not Application.Intersect(cntryRange, Range(Target.Address)) Is Nothing Then
'updates state, just in case country is filled out after state

stateRange.Value = stateRange.Value

End If

End Sub


