VERSION 5.00
Begin VB.Form frmLogin 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   2910
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4800
   ForeColor       =   &H8000000C&
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2910
   ScaleWidth      =   4800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   415
      Left            =   3000
      Picture         =   "frmLogin.frx":14432
      TabIndex        =   3
      Top             =   2280
      Width           =   1455
   End
   Begin VB.CommandButton cmdLogin 
      Caption         =   "Login"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   415
      Left            =   360
      Picture         =   "frmLogin.frx":149BC
      TabIndex        =   2
      Top             =   2280
      Width           =   1455
   End
   Begin VB.TextBox txtPwd 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   405
      IMEMode         =   3  'DISABLE
      Left            =   1200
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1440
      Width           =   3375
   End
   Begin VB.TextBox txtUsr 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   435
      Left            =   1200
      TabIndex        =   0
      Top             =   840
      Width           =   3375
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   4680
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   120
      Picture         =   "frmLogin.frx":14F46
      Top             =   120
      Width           =   480
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Type Username and Password"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   840
      TabIndex        =   6
      Top             =   240
      Width           =   3135
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H8000000C&
      BackStyle       =   1  'Opaque
      Height          =   735
      Left            =   0
      Top             =   0
      Width           =   5295
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Username:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   1095
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdLogin_Click()
Dim uSr, pWd As String
uSr = txtUsr.Text
pWd = txtPwd.Text

rs.Open "Select*From tblEmployee where Employee_ID='" & uSr & "'", cn, 3, 3
If rs.RecordCount > 0 Then
    Do Until rs.EOF
        If uSr = rs!Employee_ID And pWd = rs!Password Then
            Set rs = Nothing
                With MDIForm1
                    .lot.Enabled = True
                    .loc.Enabled = True
                    .trans.Enabled = True
                    .maint.Enabled = True
                    .rep.Enabled = True
                    .lin.Enabled = False
                    .Toolbar1.Buttons(1).Enabled = True
                    .Toolbar1.Buttons(3).Enabled = True
                    .Toolbar1.Buttons(5).Enabled = True
                    .Toolbar1.Buttons(7).Enabled = True
                    .Toolbar1.Buttons(9).Enabled = True
                    .Toolbar1.Buttons(13).Enabled = True
                   ' .Show
                End With
                ctr = 0
                Unload Me
                   Exit Do
        Else
            rs.MoveNext
            ctr = ctr + 1
        End If
    Loop
    If ctr > 0 Then
        Set rs = Nothing
        MsgBox "Access Denied!!", vbCritical, "Inventory System"
        i_Clear.cLearMe Me
        txtUsr.SetFocus
        ctr = 0
    End If
End If
End Sub

Private Sub Command1_Click()
txtUsr.Text = ""
txtPwd.Text = ""
txtUsr.SetFocus
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Me
End Sub

Private Sub txtPwd_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
cmdLogin_Click
'cmdAccess_Click
End If
End Sub
