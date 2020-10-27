VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True


Private Sub Workbook_BeforeClose(Cancel As Boolean)
Call BackToMain
End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    If InStr(UCase(Application.userName), "ENGLAND") > 0 Then
        Debug.Print "yay!"
        Call ExportModules
    End If
End Sub

Private Sub Workbook_Open()
Range("A1").Select
End Sub



