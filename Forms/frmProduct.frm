VERSION 5.00
Begin VB.Form frmProducts 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Product Information"
   ClientHeight    =   8625
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12915
   Icon            =   "frmProduct.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmProduct.frx":617A
   ScaleHeight     =   8625
   ScaleWidth      =   12915
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   120
      Picture         =   "frmProduct.frx":C2F4
      ScaleHeight     =   735
      ScaleWidth      =   855
      TabIndex        =   22
      Top             =   120
      Width           =   855
   End
   Begin VB.ListBox lstProduct 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7185
      Left            =   120
      TabIndex        =   21
      Top             =   960
      Width           =   3255
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   3840
      TabIndex        =   12
      Top             =   6960
      Width           =   8895
      Begin VB.CommandButton cmdQuit 
         Caption         =   "Quit"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   7560
         Picture         =   "frmProduct.frx":1246E
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   6360
         Picture         =   "frmProduct.frx":129F8
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "Find"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   5160
         Picture         =   "frmProduct.frx":12F82
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Delete"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   3960
         Picture         =   "frmProduct.frx":130E9
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "Update"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   2760
         Picture         =   "frmProduct.frx":13673
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdAppend 
         Caption         =   "Save"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   1560
         Picture         =   "frmProduct.frx":13BFD
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdNew 
         Caption         =   "New"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   360
         Picture         =   "frmProduct.frx":14187
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   240
         Width           =   975
      End
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
      Height          =   5775
      Left            =   3840
      TabIndex        =   0
      Top             =   960
      Width           =   8895
      Begin VB.ComboBox cboSupplier 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2040
         TabIndex        =   20
         Top             =   1680
         Width           =   4215
      End
      Begin VB.TextBox txtINStock 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2040
         TabIndex        =   11
         Top             =   3360
         Width           =   3615
      End
      Begin VB.TextBox txtUnit_Price 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2040
         TabIndex        =   10
         Top             =   2640
         Width           =   3615
      End
      Begin VB.ComboBox cboCategory 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2040
         TabIndex        =   9
         Top             =   2160
         Width           =   3135
      End
      Begin VB.TextBox txtProduct_Name 
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
         Height          =   615
         Left            =   2040
         TabIndex        =   8
         Top             =   960
         Width           =   5655
      End
      Begin VB.TextBox txtProduct_ID 
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
         Height          =   615
         Left            =   2040
         TabIndex        =   7
         Top             =   240
         Width           =   5655
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Unit_In_Stock"
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
         Left            =   120
         TabIndex        =   6
         Top             =   3480
         Width           =   1815
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Unit_Price"
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
         Left            =   120
         TabIndex        =   5
         Top             =   2880
         Width           =   1695
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Category"
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
         Left            =   120
         TabIndex        =   4
         Top             =   2160
         Width           =   1695
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Product_Name"
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
         Left            =   120
         TabIndex        =   3
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Supplier"
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
         Left            =   120
         TabIndex        =   2
         Top             =   1680
         Width           =   1695
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Product_ID"
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
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Product Record"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   495
      Left            =   1320
      TabIndex        =   23
      Top             =   240
      Width           =   4335
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000005&
      Height          =   855
      Left            =   0
      Top             =   0
      Width           =   12735
   End
End
Attribute VB_Name = "frmProducts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Text1_Change()

End Sub

Private Sub cmdAppend_Click()
Dim pIDCheck$
pIDCheck = txtProduct_ID

Field_Check.Empty_Checks Me
If iTerminate = True Then
    iTerminate = False
    Exit Sub
End If

If ctrl_Flag = False Then
    Call Redundancy_Check(pIDCheck)
    If iTerminate = True Then
        MsgBox "The product you are trying to save is" & vbCrLf & "already on the record!", vbExclamation, "Inventory System"
        iTerminate = False
        Exit Sub
    End If
End If

prod_ID = txtProduct_ID.Text
Call SQL_Execute(ctrl_Flag, prod_ID)
i_Clear.cLearMe Me
i_Disable.Disable_Txt Me
cmdAppend.Enabled = False
cmdUpdate.Enabled = True
ctrl_Flag = False
prod_ID = ""

End Sub


Private Sub cmdCancel_Click()
ctrl_Flag = False
i_Clear.cLearMe Me
i_Disable.Disable_Txt Me
cmdAppend.Enabled = False
cmdUpdate.Enabled = True
cmdNew.Enabled = True
End Sub

Private Sub cmdDelete_Click()
i_Delete = InputBox("Enter Stock code of an item to delete.")
rs.Open "select*from tblProduct where Product_ID='" & i_Delete & "'", cn, 3, 3
rs.Delete
Set rs = Nothing
Call Reload_List
End Sub

Private Sub cmdFind_Click()
Dim pID As String
pID = InputBox("Please enter product ID to search")
Call Srch_Rec(Me.Name, pID)
End Sub

Private Sub cmdNew_Click()

cmdAppend.Enabled = True
ctrl_Flag = False
i_Enable.Enable_Txt Me
i_Clear.cLearMe Me
txtProduct_ID.SetFocus
cmdNew.Enabled = False
End Sub

Private Sub cmdQuit_Click()
Unload Me
End Sub

Private Sub cmdUpdate_Click()
ctrl_Flag = True
prod_ID = InputBox("Please enter product ID to update.")
If prod_ID = "" Then
    ctrl_Flag = False
    Exit Sub
Else
    Call Srch_Rec(Me.Name, prod_ID)
    i_Enable.Enable_Txt Me
    txtProduct_ID.Locked = True
    txtProduct_ID.SetFocus
    cmdAppend.Enabled = True
    cmdUpdate.Enabled = False
End If
End Sub

Private Sub Form_Load()
rs.Open "select*from tblProduct", cn, 3, 3
Do While Not rs.EOF
    lstProduct.AddItem rs!Product_Name
    rs.MoveNext
Loop
Set rs = Nothing

'Collects the Category from the table
rs.Open "Select distinct Category from tblProduct", cn, 3, 3
If rs.RecordCount > 0 Then
    Do While Not rs.EOF
        cboCategory.AddItem rs!Category
        rs.MoveNext
    Loop
End If
Set rs = Nothing

rs.Open "Select*from tblSupplier", cn, 3, 3
If rs.RecordCount > 0 Then
    Do While Not rs.EOF
        cboSupplier.AddItem rs!supplier_name
        rs.MoveNext
    Loop
End If
Set rs = Nothing

End Sub


'Redundancy checking function
Private Function Redundancy_Check(pIDCheck)
rs.Open "Select*from tblProduct where Product_ID='" & pIDCheck & "'", cn, 3, 3
If rs.RecordCount > 0 Then
    If Not (rs.BOF And rs.EOF) Then
        iTerminate = True
    End If
End If
Set rs = Nothing
End Function






Private Sub lstProduct_Click()
Dim pNam As String
If lstProduct.ListCount > 0 Then
    pNam = lstProduct.Text
    rs.Open "Select*from tblProduct where Product_Name='" & pNam & "'", cn, 3, 3
    txtProduct_ID.Text = rs!Product_ID
    txtProduct_Name.Text = rs!Product_Name
    cboSupplier.Text = rs!Supplier
    cboCategory.Text = rs!Category
    txtUnit_Price.Text = rs!Unit_Price
    txtINStock.Text = rs!Unit_In_Stock
    Set rs = Nothing
Else
     Exit Sub
End If
End Sub






