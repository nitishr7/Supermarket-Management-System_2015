VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmSupplier 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Supplier Profile"
   ClientHeight    =   8625
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11820
   Icon            =   "frmSupplier.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8625
   ScaleWidth      =   11820
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid supGrid 
      Height          =   3255
      Left            =   120
      TabIndex        =   11
      Top             =   5040
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   5741
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   11535
      Begin VB.TextBox txtSupID 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3120
         TabIndex        =   6
         Top             =   240
         Width           =   4575
      End
      Begin VB.TextBox txtSupName 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3120
         TabIndex        =   5
         Top             =   720
         Width           =   4575
      End
      Begin VB.TextBox txtAddress 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3120
         TabIndex        =   4
         Top             =   1200
         Width           =   4575
      End
      Begin VB.TextBox txtTelephone 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3120
         TabIndex        =   3
         Top             =   1680
         Width           =   4575
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Supplier ID"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   10
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Supplier Name"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Telephone"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   1680
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   3600
      Width           =   11535
      Begin VB.CommandButton cmdQuit 
         Caption         =   "Quit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   7560
         Picture         =   "frmSupplier.frx":1D8A
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   6360
         Picture         =   "frmSupplier.frx":2314
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "Find"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   5160
         Picture         =   "frmSupplier.frx":289E
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Delete"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   3960
         Picture         =   "frmSupplier.frx":2A05
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "Update"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   2760
         Picture         =   "frmSupplier.frx":2F8F
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdAppend 
         Caption         =   "Save"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   1560
         Picture         =   "frmSupplier.frx":3519
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdNew 
         Caption         =   "New"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   360
         Picture         =   "frmSupplier.frx":3AA3
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Suppliers Record"
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
      TabIndex        =   19
      Top             =   120
      Width           =   4335
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Type supplier information and then click Save"
      Height          =   375
      Left            =   1080
      TabIndex        =   18
      Top             =   600
      Width           =   4335
   End
   Begin VB.Image Image1 
      Height          =   855
      Left            =   120
      Picture         =   "frmSupplier.frx":402D
      Top             =   120
      Width           =   810
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000E&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000004&
      Height          =   1095
      Left            =   -120
      Top             =   0
      Width           =   11775
   End
End
Attribute VB_Name = "frmSupplier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAppend_Click()
Field_Check.Empty_Checks Me
If iTerminate = True Then
    iTerminate = False
    Exit Sub
End If


Sup_ID = UCase(txtSupID.Text)
Sup_Name = txtSupName.Text
sAdd = txtAddress.Text
sTel = txtTelephone.Text

If ctrl_Flag = False Then
    Call Redundancy_Check(Sup_ID)
    If iTerminate = True Then
        MsgBox "The supplier you are trying to save is" & vbCrLf & "already on the record!", vbExclamation, "Inventory System"
        iTerminate = False
        Exit Sub
    End If
End If


Select Case ctrl_Flag
    Case False:
        rs.Open "Insert Into tblSupplier(Supplier_ID,Supplier_Name,Address,Telephone)" & _
        "Values('" & Sup_ID & "','" & Sup_Name & "','" & sAdd & "','" & sTel & "')", cn, 3, 3
        Set rs = Nothing
        i_Clear.cLearMe Me
        i_Disable.Disable_Txt Me
        cmdAppend.Enabled = False
        Call grd_Data_Loader
        MsgBox "One new supplier has been saved.", vbInformation, "Inventory System"
    Case True:
        rs.Open "Update tblSupplier set Supplier_Name='" & Sup_Name & "',Address='" & sAdd & "',Telephone='" & sTel & "'" & _
        "Where Supplier_ID='" & Sup_ID & "'", cn, 3, 3
        i_Disable.Disable_Txt Me
        Set rs = Nothing
        Call grd_Data_Loader
        MsgBox "Supplier with supplier ID " & Sup_ID & " has been updated.", vbInformation, "Inventory System"
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
txtSupID.Locked = False
cmdUpdate.Enabled = True
cmdNew.Enabled = True
End Sub

Private Sub cmdDelete_Click()
iFind = UCase(InputBox("Enter supplier ID to delete."))
Call sup_Find
If ctr = 0 Then
    rs.Open "Delete*from tblSupplier where Supplier_ID='" & UCase(iFind) & "'", cn, 3, 3
    Set rs = Nothing
    Call grd_Data_Loader
    i_Clear.cLearMe Me
    MsgBox "Supplier with supplier ID " & iFind & " has been deleted.", vbInformation, "Inventory System"
End If
End Sub

Private Sub cmdFind_Click()
iFind = UCase(InputBox("Enter supplier ID to find."))
Call sup_Find
End Sub

Private Sub cmdNew_Click()
i_Enable.Enable_Txt Me
i_Clear.cLearMe Me
txtSupID.SetFocus
cmdAppend.Enabled = True
txtSupID.Locked = False
ctrl_Flag = False
cmdNew.Enabled = False
End Sub

Private Sub MSHFlexGrid1_Click()

End Sub

Private Function grd_Data_Loader()
rs.Open "Select*from tblSupplier", cn, 3, 3
Set supGrid.DataSource = rs
Set rs = Nothing
End Function

Private Sub cmdQuit_Click()
Unload Me
End Sub

Private Sub cmdUpdate_Click()
iFind = UCase(InputBox("Enter supplier ID to update."))
Call sup_Find
If ctr = 0 Then
    ctrl_Flag = True
    i_Enable.Enable_Txt Me
    txtSupID.Locked = True
    cmdAppend.Enabled = True
    cmdUpdate.Enabled = False
End If
End Sub

Private Sub Form_Load()
With supGrid
    .ColWidth(0) = 300
    .ColWidth(1) = 1500
    .ColWidth(2) = 3000
    .ColWidth(3) = 4000
    .ColWidth(4) = 1400
End With
Call grd_Data_Loader
End Sub



'Redundancy checking function
Private Function Redundancy_Check(Sup_ID)
rs.Open "Select*from tblSupplier where Supplier_ID='" & Sup_ID & "'", cn, 3, 3
If rs.RecordCount > 0 Then
    If Not (rs.BOF And rs.EOF) Then
        iTerminate = True
    End If
End If
Set rs = Nothing
End Function




Private Sub supGrid_Click()
X = supGrid.Row
With supGrid
    txtSupID.Text = .TextMatrix(X, 1)
    txtSupName.Text = .TextMatrix(X, 2)
    txtAddress.Text = .TextMatrix(X, 3)
    txtTelephone.Text = .TextMatrix(X, 4)
End With









End Sub
