VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form frmPOrder 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Purchase Order"
   ClientHeight    =   7860
   ClientLeft      =   4230
   ClientTop       =   2880
   ClientWidth     =   8790
   Icon            =   "frmPO.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7860
   ScaleWidth      =   8790
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Product List"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   120
      TabIndex        =   14
      Top             =   2280
      Width           =   8535
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdProduct 
         Height          =   1695
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   2990
         _Version        =   393216
         Enabled         =   0   'False
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Order List"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   120
      TabIndex        =   12
      Top             =   5040
      Width           =   8535
      Begin MSComctlLib.ListView lvOrderedList 
         Height          =   1695
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   2990
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Product_ID"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Product_Name"
            Object.Width           =   5821
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "QTY"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Supplier"
            Object.Width           =   7056
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Category"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Unit_Price"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Unit_In_Stock"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "Remove"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7560
      Picture         =   "frmPO.frx":617A
      TabIndex        =   11
      Top             =   4440
      Width           =   1095
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6240
      Picture         =   "frmPO.frx":6704
      TabIndex        =   10
      Top             =   4440
      Width           =   1215
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7320
      Picture         =   "frmPO.frx":6C8E
      TabIndex        =   9
      Top             =   7200
      Width           =   1335
   End
   Begin VB.CommandButton cmdPurchase 
      Caption         =   "Purchased"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5880
      Picture         =   "frmPO.frx":7218
      TabIndex        =   8
      Top             =   7200
      Width           =   1335
   End
   Begin MSComCtl2.DTPicker DPDate_Required 
      Height          =   375
      Left            =   6600
      TabIndex        =   3
      Top             =   1680
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      Format          =   87425025
      CurrentDate     =   38753
   End
   Begin MSComCtl2.DTPicker DPDate_Order 
      Height          =   375
      Left            =   6600
      TabIndex        =   2
      Top             =   1200
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      Format          =   87425025
      CurrentDate     =   38753
   End
   Begin VB.ComboBox cboSupplier 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1320
      TabIndex        =   1
      Top             =   1800
      Width           =   3015
   End
   Begin VB.TextBox txtPO 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1320
      TabIndex        =   0
      Top             =   1320
      Width           =   3015
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Purchase Order"
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
      TabIndex        =   20
      Top             =   360
      Width           =   4335
   End
   Begin VB.Image Image1 
      Height          =   855
      Left            =   120
      Picture         =   "frmPO.frx":77A2
      Top             =   120
      Width           =   810
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Item"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   7200
      Width           =   1215
   End
   Begin VB.Label lblstat 
      BackStyle       =   0  'Transparent
      Caption         =   "item"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   255
      Left            =   1560
      TabIndex        =   18
      Top             =   7200
      Width           =   4695
   End
   Begin VB.Label lblItemName 
      BackStyle       =   0  'Transparent
      Caption         =   "item"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1560
      TabIndex        =   17
      Top             =   4440
      Width           =   4575
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Item Name:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   16
      Top             =   4440
      Width           =   1215
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Supplier:"
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
      TabIndex        =   7
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "PO No:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Date Order:"
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
      Left            =   5160
      TabIndex        =   5
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Date Required:"
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
      Left            =   5160
      TabIndex        =   4
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000E&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000004&
      Height          =   1095
      Left            =   0
      Top             =   0
      Width           =   10575
   End
End
Attribute VB_Name = "frmPOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cboSupplier_Click()
rs.Open "select*from tblProduct where Supplier='" & cboSupplier.Text & "'", cn, 3, 3
If Not (rs.BOF And rs.EOF) Then grdProduct.Enabled = True
Set grdProduct.DataSource = rs
Set rs = Nothing
End Sub

Private Sub cmdAdd_Click()
With frmAdd
    .lblProd_ID.Text = scAd
    .lblProd_Name.Text = itmAd
End With
frmAdd.Show vbModal
cmdPurchase.Enabled = True
cmdAdd.Enabled = False
    lblstat.Caption = lvOrderedList.ListItems.Count
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdPurchase_Click()
Dim sup, DOr, DRe, SC, QTY As String

poN = UCase(txtPO.Text)
sup = cboSupplier.Text
DOr = DPDate_Order.Value
DRe = DPDate_Required.Value

'Checks for empty fields
If lvOrderedList.ListItems.Count = 0 Then MsgBox "There is no item to be save", vbExclamation, "Empty list of Item": Exit Sub
Field_Check.Empty_Checks Me
If iTerminate = True Then
    iTerminate = False
    Exit Sub
End If

'Checks for duplicate
Call PO_Duplicate_Check(poN)
If iTerminate = True Then
    MsgBox "The PO Number you are trying to purchase " & vbCrLf & "is already exist!", vbExclamation, "Inventory System"
    iTerminate = False
    Exit Sub
End If


'Saves data file by batch to PO table
For i = 1 To lvOrderedList.ListItems.Count
    SC = lvOrderedList.ListItems(i).Text
    QTY = lvOrderedList.ListItems(i).ListSubItems(2).Text
    cn.Execute "Insert Into tblPO(PO_No,Product_ID,Quantity,PO_Supplier,PO_Date_Order,PO_Date_Required)" & _
    "Values('" & poN & "','" & SC & "','" & QTY & "','" & sup & "','" & DOr & "','" & DRe & "')"
Next
lvOrderedList.ListItems.Clear
cmdPurchase.Enabled = False
cmdAdd.Enabled = False
MsgBox "Purchase transaction No " & poN & " has been purchased.", vbInformation, "Inventory system"
Call PO_AutoNum
End Sub

Private Sub cmdRemove_Click()
Dim i As Integer
For i = lvOrderedList.ListItems.Count To 1 Step -1
    If lvOrderedList.ListItems(i).Checked = True Then
        lvOrderedList.ListItems.Remove i
    End If
    lblstat.Caption = lvOrderedList.ListItems.Count
Next i

End Sub

Private Sub Form_Load()
With grdProduct
    .ColWidth(0) = 300
    .ColWidth(1) = 1500
    .ColWidth(2) = 2500
    .ColWidth(3) = 2500
    .ColWidth(6) = 1400
End With

Call PO_AutoNum
With grdProduct
    .ColWidth(0) = 200
'    .TextMatrix(1, 1) = "Fuck" '(1,=row 1=col)
End With

Set rs = Nothing


rs.Open "Select*from tblSupplier", cn, 3, 3
If rs.RecordCount > 0 Then
    Do Until rs.EOF
        cboSupplier.AddItem rs!supplier_name
        rs.MoveNext
    Loop
End If
Set rs = Nothing
lblstat.Caption = lvOrderedList.ListItems.Count

With grdProduct
    
End With
End Sub

Private Sub grdProduct_Click()
i = grdProduct.Row
With grdProduct
    lblItemName.Caption = .TextMatrix(i, 2)
    scAd = .TextMatrix(i, 1)
    itmAd = .TextMatrix(i, 2)
End With
If Not scAd = "" Then
    cmdAdd.Enabled = True
End If
End Sub

'Checks for PO redundancy occurences
Private Function PO_Duplicate_Check(poN)
rs.Open "Select*from tblPO where PO_No='" & poN & "'", cn, 3, 3
If rs.RecordCount > 0 Then
    If Not (rs.BOF And rs.EOF) Then
        iTerminate = True
    End If
End If
Set rs = Nothing

End Function



Private Function PO_AutoNum()
rs.Open "select * from tblPO Order By PO_No DESC", cn, 3, 2
If rs.RecordCount = 0 Then
    txtPO.Text = "PO-0000"
Else
    txtPO.Text = "PO-" & Format(Right(rs!PO_NO, 4) + 1, "0000")
End If
rs.Close
txtPO.Locked = True
End Function

