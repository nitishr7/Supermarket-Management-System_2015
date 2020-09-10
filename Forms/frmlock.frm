VERSION 5.00
Begin VB.Form frmlock 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "System Lock"
   ClientHeight    =   1035
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   4125
   Icon            =   "frmlock.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1035
   ScaleWidth      =   4125
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Log-out"
      Height          =   375
      Left            =   2640
      TabIndex        =   3
      Top             =   600
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Enter"
      Height          =   375
      Left            =   2400
      TabIndex        =   2
      Top             =   1800
      Width           =   735
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1440
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   240
      Width           =   2535
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
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
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "frmlock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim pWd As String
pWd = Text1.Text

rs.Open "Select*From tblEmployee", cn, 3, 3
If rs.RecordCount > 0 Then
    Do Until rs.EOF
        If pWd = rs!Password Then
            Set rs = Nothing
                With MDIForm1
                    .lot.Enabled = True
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
                End With
                ctr = 0
                Unload Me
                MDIForm1.Enabled = True
                MDIForm1.Show
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
        Text1.SetFocus
        ctr = 0
    End If
End If
End Sub

Private Sub Command2_Click()
With MDIForm1
.lin.Enabled = True
.trans.Enabled = False
.maint.Enabled = False
.rep.Enabled = False
.lot.Enabled = False
.Toolbar1.Buttons(1).Enabled = False
.Toolbar1.Buttons(3).Enabled = False
.Toolbar1.Buttons(5).Enabled = False
.Toolbar1.Buttons(7).Enabled = False
.Toolbar1.Buttons(9).Enabled = False
.Toolbar1.Buttons(13).Enabled = False
End With
Unload Me
MDIForm1.Enabled = True
MDIForm1.Show
End Sub

Private Sub Form_Load()
MDIForm1.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
MDIForm1.Enabled = True
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Command1_Click
'cmdAccess_Click
End If
End Sub
