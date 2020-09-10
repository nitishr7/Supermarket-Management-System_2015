VERSION 5.00
Begin VB.Form frmAdminAutorize 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Admin Authorization"
   ClientHeight    =   600
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4545
   Icon            =   "frmAdminAutorize.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   600
   ScaleWidth      =   4545
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Height          =   415
      Left            =   3000
      Picture         =   "frmAdminAutorize.frx":1D8A
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2280
      Width           =   735
   End
   Begin VB.TextBox txtAdminPwd 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      IMEMode         =   3  'DISABLE
      Left            =   1800
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   150
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Password"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   1815
   End
End
Attribute VB_Name = "frmAdminAutorize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()
ctr = 0
If rs.State = 1 Then Set rs = Nothing
rs.Open "Select*from tblEmployee", cn, 3, 3
If rs.RecordCount > 0 Then
    Do Until rs.EOF
        If rs!designation = "Administrator" And rs!Password = txtAdminPwd.Text Then
            Set rs = Nothing
            Unload Me
            frmUser.Show vbModal
            ctr = 0
            Exit Do
        Else
            rs.MoveNext
            ctr = ctr + 1
        End If
    Loop
    If ctr > 0 Then
        Set rs = Nothing
        MsgBox "Incorrect pasword. You are not authorized" & vbCrLf & "to modify any user account!", vbCritical, "Inventory system"
        ctr = 0
        txtAdminPwd.Text = ""
        txtAdminPwd.SetFocus
    End If
End If
End Sub

Private Sub txtAdminPwd_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
cmdOK_Click
'cmdAccess_Click
End If
End Sub
