VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Activate()

ThisWorkbook.RefreshAll
Call CheckDB

End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
'checks DB when DB info is changed

Dim dbLoc As Range, dbName As Range, performCheck As Boolean

Set dbLoc = Range("A" & Application.WorksheetFunction.Match("D*B*Loc*", Range("A:A"), 0) + 1) 'DB location
Set dbName = Range("A" & Application.WorksheetFunction.Match("D*B*Name*", Range("A:A"), 0) + 1) 'DB filename
performCheck = False

If Not Application.Intersect(dbLoc, Range(Target.Address)) Is Nothing Then
    
    performCheck = True
    If Right(dbLoc.Value, 1) <> "\" Then 'needed for database query
        dbLoc.Value = dbLoc.Value & "\"
        performCheck = False 'changing cell value would give double error msg
    End If
    
ElseIf Not Application.Intersect(dbName, Range(Target.Address)) Is Nothing Then
    
    performCheck = True
    If Right(dbName.Value, 6) <> ".accdb" Then 'needed for database query
        dbName.Value = dbName.Value & ".accdb"
        performCheck = False 'changing cell value would give double error msg
    End If
    
End If

If performCheck Then
    Call CheckDB
End If

End Sub
