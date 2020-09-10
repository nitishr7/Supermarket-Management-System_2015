VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TextHandling"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Function cLearMe(ByRef frm As Form) As String
Dim Form_Object As Control
For Each Form_Object In frm
    If TypeOf Form_Object Is TextBox Or TypeOf Form_Object Is ComboBox Then
        Form_Object.Text = ""
    End If
Next Form_Object
End Function

Function Enable_Txt(ByRef frm As Form) As String
Dim Form_Object As Control
For Each Form_Object In frm
    If TypeOf Form_Object Is TextBox Or TypeOf Form_Object Is ComboBox Then
        Form_Object.Enabled = True
    End If
Next Form_Object
End Function

Function Empty_Checks(ByRef frm As Form) As String
Dim Form_Object As Control
For Each Form_Object In frm
    If TypeOf Form_Object Is TextBox Or TypeOf Form_Object Is ComboBox Then
        If Form_Object.Text = "" Then
            MsgBox Form_Object.Name & " is empty. Please complete the field.", vbExclamation
            iTerminate = True
            Form_Object.SetFocus
            Exit Function
        End If
    End If
Next Form_Object
iTerminate = False
End Function

Function Disable_Txt(ByRef frm As Form) As String
Dim Form_Object As Control
For Each Form_Object In frm
    If TypeOf Form_Object Is TextBox Or TypeOf Form_Object Is ComboBox Then
        Form_Object.Enabled = False
    End If
Next Form_Object
End Function



