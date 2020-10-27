VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub importAccessData()

Dim accConn As ADODB.Connection, recSet As ADODB.Recordset, sQRY As String
Dim dbPath As String, dbFile As String, dbFileName As String, tableName As String

On Error GoTo errhandler

dbPath = Sheet1.Range("A" & Application.WorksheetFunction.Match("D*B*Loc*", Sheet1.Range("A:A"), 0) + 1).Value
dbFileName = Sheet1.Range("A" & Application.WorksheetFunction.Match("D*B*Name*", Sheet1.Range("A:A"), 0) + 1).Value

MsgBox dbPath & ", " & dbFileName

If Right(dbPath, 1) <> "\" Then
    dbPath = dbPath & "\"
End If

dbFileName = Dir(dbPath & dbFileName & "*")

On Error Resume Next
    Do While Right(dbFileName, 6) <> ".accdb"
        dbFileName = Dir()
        If dbFileName = "" Then
            dbFileName = ".accdb"
        End If
    Loop
On Error GoTo errhandler

If dbFileName = ".accdb" Then
    MsgBox "Access-linked table could not be updated."
    Exit Sub
End If

dbFile = dbPath & dbFileName
tableName = ""

Set accConn = New ADODB.Connection

Set recSet = New ADODB.Recordset

accConn.Open "Provider=Microsoft.ACE.OLEDB.12.0;" & _
"Data Source=" & dbFile & ";"

sQRY = "SELECT * FROM " & tableName

recSet.CursorLocation = adUseClient

recSet.Open sQRY, accConn, adOpenStatic, adLockReadOnly

Application.ScreenUpdating = False

Range(“A50”).CopyFromRecordset recSet

recSet.Close

Set recSet = Nothing

accConn.Close

Set accConn = Nothing

Exit Sub

errhandler:
MsgBox "Failure"

End Sub


Private Sub Worksheet_Activate()

ThisWorkbook.RefreshAll
'Call UpdateCustomerDropdown

End Sub

