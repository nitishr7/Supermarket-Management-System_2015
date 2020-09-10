VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmUser 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "User Profile (Administrator full access)"
   ClientHeight    =   4800
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6375
   Icon            =   "frmUser.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4800
   ScaleWidth      =   6375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtUsrID 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1560
      TabIndex        =   15
      Top             =   1320
      Width           =   2295
   End
   Begin VB.TextBox txtFulName 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1560
      TabIndex        =   14
      Top             =   1800
      Width           =   3855
   End
   Begin VB.TextBox txtDepartment 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1560
      TabIndex        =   13
      Top             =   2760
      Width           =   2295
   End
   Begin VB.TextBox txtPwd 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1560
      PasswordChar    =   "*"
      TabIndex        =   12
      Top             =   3240
      Width           =   3135
   End
   Begin VB.ComboBox txtDesignation 
      Height          =   315
      ItemData        =   "frmUser.frx":1D8A
      Left            =   1560
      List            =   "frmUser.frx":1D94
      TabIndex        =   11
      Text            =   "Combo1"
      Top             =   2280
      Width           =   2295
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   3720
      Width           =   6135
      Begin VB.CommandButton cmdQuit 
         Caption         =   "Quit"
         Height          =   615
         Left            =   5160
         Picture         =   "frmUser.frx":1DAE
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   240
         Width           =   865
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   615
         Left            =   4320
         Picture         =   "frmUser.frx":2338
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   865
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "Find"
         Height          =   615
         Left            =   3480
         Picture         =   "frmUser.frx":28C2
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   240
         Width           =   865
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Delete"
         Height          =   615
         Left            =   2640
         Picture         =   "frmUser.frx":2A29
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   240
         Width           =   865
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "Update"
         Height          =   615
         Left            =   1800
         Picture         =   "frmUser.frx":2FB3
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Width           =   865
      End
      Begin VB.CommandButton cmdAppend 
         Caption         =   "Save"
         Enabled         =   0   'False
         Height          =   615
         Left            =   960
         Picture         =   "frmUser.frx":353D
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   865
      End
      Begin VB.CommandButton cmdNew 
         Caption         =   "New"
         Height          =   615
         Left            =   120
         Picture         =   "frmUser.frx":3AC7
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Width           =   865
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid usrGrid 
      Height          =   1935
      Left            =   120
      TabIndex        =   8
      Top             =   5280
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   3413
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "User ID"
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
      Left            =   360
      TabIndex        =   20
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Designation"
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
      Left            =   360
      TabIndex        =   19
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Full Name"
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
      Left            =   360
      TabIndex        =   18
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Department"
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
      Left            =   360
      TabIndex        =   17
      Top             =   2760
      Width           =   1695
   End
   Begin VB.Label Label5 
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
      Left            =   360
      TabIndex        =   16
      Top             =   3240
      Width           =   1695
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "User Management"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   495
      Left            =   1080
      TabIndex        =   10
      Top             =   120
      Width           =   4335
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Username and Password"
      Height          =   375
      Left            =   1080
      TabIndex        =   9
      Top             =   600
      Width           =   4335
   End
   Begin VB.Image Image1 
      Height          =   855
      Left            =   120
      Picture         =   "frmUser.frx":4051
      Top             =   120
      Width           =   810
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000E&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000004&
      Height          =   1095
      Left            =   0
      Top             =   0
      Width           =   7095
   End
End
Attribute VB_Name = "frmUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdAppend_Click()
Dim usrID, uName, uDes, uDep, uPwd As String
'Checks for empty fields
Field_Check.Empty_Checks Me
If iTerminate = True Then
    iTerminate = False
    Exit Sub
End If

'Assigns the variable values
usrID = txtUsrID.Text
uName = txtFulName.Text
uDes = txtDesignation.Text
uDep = txtDepartment.Text
uPwd = txtPwd.Text
If ctrl_Flag = False Then
Call Redundancy_Check(usrID)
    If iTerminate = True Then
        MsgBox "The user you are trying to save is" & vbCrLf & "already on the record!", vbExclamation, "Inventory System"
        iTerminate = False
        Exit Sub
    End If
End If

'Saves in case of boolean expressions
Select Case ctrl_Flag
    Case False:
        rs.Open "Insert Into tblEmployee(Employee_ID,Employee_Name,Designation,Department,Password)" & _
        "Values('" & usrID & "','" & uName & "','" & uDes & "','" & uDep & "','" & uPwd & "')", cn, 3, 3
        Set rs = Nothing
        Call usrGrid_Load
        i_Clear.cLearMe Me
        i_Disable.Disable_Txt Me
        MsgBox "New user has been saved.", vbInformation, "Inventory System"
    Case True:
        rs.Open "Update tblEmployee set Employee_name='" & uName & "',Designation='" & uDes & "',Department='" & uDep & "',Password='" & uPwd & "'" & _
        "Where Employee_ID='" & iFind & "'", cn, 3, 3
        Set rs = Nothing
        Call usrGrid_Load
        i_Disable.Disable_Txt Me
        MsgBox "User account has been updated.", vbInformation, "Inventory System"
End Select
ctrl_Flag = False
cmdAppend.Enabled = False
cmdUpdate.Enabled = True
End Sub

Private Sub cmdCancel_Click()
ctrl_Flag = False
cmdAppend.Enabled = False
i_Clear.cLearMe Me
i_Disable.Disable_Txt Me
txtUsrID.Locked = False
cmdUpdate.Enabled = True
cmdNew.Enabled = True
End Sub

Private Sub cmdDelete_Click()
iFind = InputBox("Enter user ID to delete.")
Call usr_Find
If ctr = 0 Then
    If iFind = "Administrator" Then
        i_Clear.cLearMe Me
        MsgBox "Deleting Administrator account is not allowed!", vbCritical, "Inventory System"
        Exit Sub
    Else
        rs.Open "Delete*from tblEmployee where Employee_ID='" & iFind & "'", cn, 3, 3
        Set rs = Nothing
        Call usrGrid_Load
        i_Clear.cLearMe Me
        MsgBox "User with User ID " & iFind & " has been deleted.", vbInformation, "Inventory System"
    End If
End If

End Sub

Private Sub cmdFind_Click()
iFind = InputBox("Enter user ID to find.")
Call usr_Find
End Sub

Private Sub cmdNew_Click()
i_Clear.cLearMe Me
i_Enable.Enable_Txt Me
ctrl_Flag = False
cmdAppend.Enabled = True
txtUsrID.SetFocus
cmdNew.Enabled = False
End Sub

Private Sub cmdQuit_Click()
Unload Me
End Sub

Private Sub cmdUpdate_Click()
iFind = InputBox("Enter user ID to update.")
Call usr_Find
If ctr = 0 Then
    ctrl_Flag = True
    cmdAppend.Enabled = True
    i_Enable.Enable_Txt Me
    txtUsrID.Locked = True
    txtFulName.SetFocus
    cmdUpdate.Enabled = False
End If

End Sub

Private Sub Form_Load()
Call usrGrid_Load
ctrl_Flag = False
cmdAppend.Enabled = False
i_Clear.cLearMe Me
i_Disable.Disable_Txt Me
txtUsrID.Locked = False
cmdUpdate.Enabled = True
cmdNew.Enabled = True
With usrGrid
    .ColWidth(0) = 200
    .ColWidth(1) = 1100
    .ColWidth(2) = 2000
End With
End Sub

Private Function usrGrid_Load()
rs.Open "Select*from tblEmployee", cn, 3, 3
Set usrGrid.DataSource = rs
Set rs = Nothing
End Function


'Redundancy checking function
Private Function Redundancy_Check(usrID)
rs.Open "Select*from tblEmployee where Employee_ID='" & usrID & "'", cn, 3, 3
If rs.RecordCount > 0 Then
    If Not (rs.BOF And rs.EOF) Then
        iTerminate = True
    End If
End If
Set rs = Nothing
End Function




Private Sub Label11_Click()
End Sub

Private Sub usrGrid_Click()
X = usrGrid.Row
With usrGrid
    txtUsrID.Text = .TextMatrix(X, 1)
    txtFulName.Text = .TextMatrix(X, 2)
    txtDesignation.Text = .TextMatrix(X, 3)
    txtDepartment.Text = .TextMatrix(X, 4)
    txtPwd.Text = .TextMatrix(X, 5)
End With
End Sub
