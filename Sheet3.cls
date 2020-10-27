VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Change(ByVal Target As Range)

Dim i As Integer, j As Integer
Dim stateRange As Range, cntryRange As Range

j = Application.WorksheetFunction.Match("Comp*Cat*", Range("5:5"), 0)
For i = 5 To 75
    If Cells(i, j).Value = 0 Then
        Cells(i, j).Value = "Complaint Categories" 'keeps comp categories from being blank
    End If
Next i

j = Application.WorksheetFunction.Match("*Supplier*", Range("5:5"), 0)
For i = 5 To 75
    If Cells(i, j).Value = 0 Then
        Cells(i, j).Value = "Suppliers" 'keeps supplier from being blank
    End If
Next i

j = Application.WorksheetFunction.Match("*Root*Cat*", Range("5:5"), 0)
For i = 5 To 75
    If Cells(i, j).Value = 0 Then
        Cells(i, j).Value = "Root Cause Categories" 'keeps RC categories from being blank
    End If
Next i

Set stateRange = Range("B" & Application.WorksheetFunction.Match("State*", Range("A:A"), 0))
Set cntryRange = Range("B" & Application.WorksheetFunction.Match("Country*", Range("A:A"), 0))

If Not Application.Intersect(stateRange, Range(Target.Address)) Is Nothing Then
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

