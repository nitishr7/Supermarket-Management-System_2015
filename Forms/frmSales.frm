VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form frmSales 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "New Sales"
   ClientHeight    =   9270
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15450
   Icon            =   "frmSales.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9270
   ScaleWidth      =   15450
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton optProdName 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Product Name"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3240
      TabIndex        =   42
      Top             =   1680
      Width           =   2055
   End
   Begin VB.OptionButton optProdID 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Product ID"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3240
      TabIndex        =   41
      Top             =   1200
      Width           =   2055
   End
   Begin VB.CommandButton cmdPayment 
      Caption         =   "Payment"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   3960
      Picture         =   "frmSales.frx":13C32
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   8040
      Width           =   2175
   End
   Begin VB.TextBox txtCash 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8400
      TabIndex        =   34
      Top             =   7440
      Width           =   2055
   End
   Begin VB.CommandButton cmdTot 
      Caption         =   "Total"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   6720
      Picture         =   "frmSales.frx":19DAC
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   8040
      Width           =   2175
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H8000000E&
      Height          =   3975
      Left            =   600
      TabIndex        =   5
      Top             =   480
      Width           =   14055
      Begin VB.CommandButton cmdSearch 
         Caption         =   "Search"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5520
         TabIndex        =   38
         Top             =   840
         Width           =   1935
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   7440
         Picture         =   "frmSales.frx":1A336
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   2400
         Width           =   1455
      End
      Begin VB.TextBox txtQuantity 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   9720
         TabIndex        =   13
         Top             =   1200
         Width           =   2535
      End
      Begin VB.TextBox txtStocks 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   9720
         TabIndex        =   12
         Top             =   720
         Width           =   2895
      End
      Begin VB.TextBox txtPrice 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   9720
         TabIndex        =   11
         Top             =   240
         Width           =   2895
      End
      Begin VB.TextBox txtCategory 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2520
         TabIndex        =   10
         Top             =   3240
         Width           =   3855
      End
      Begin VB.TextBox txtSupplier 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2520
         TabIndex        =   9
         Top             =   2760
         Width           =   3855
      End
      Begin VB.ComboBox cmbProdID 
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
         Height          =   405
         Left            =   2520
         TabIndex        =   8
         Top             =   1800
         Width           =   3855
      End
      Begin VB.ComboBox cmbProdName 
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
         Height          =   405
         Left            =   2520
         TabIndex        =   7
         Top             =   2280
         Width           =   3855
      End
      Begin VB.CommandButton cmdRemove 
         Caption         =   "Remove"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   9480
         Picture         =   "frmSales.frx":1A8C0
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   2400
         Width           =   1455
      End
      Begin MSComCtl2.DTPicker DPDate_Released 
         Height          =   375
         Left            =   9720
         TabIndex        =   15
         Top             =   1680
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   87752705
         CurrentDate     =   38753
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "Search by"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   480
         TabIndex        =   40
         Top             =   840
         Width           =   3495
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Transaction No."
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
         Left            =   360
         TabIndex        =   27
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label lblSO 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   2520
         TabIndex        =   25
         Top             =   120
         Width           =   3855
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Quantity"
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
         Left            =   8160
         TabIndex        =   23
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Stock(s)"
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
         Left            =   8160
         TabIndex        =   22
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Price"
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
         Left            =   7920
         TabIndex        =   21
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
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
         Left            =   600
         TabIndex        =   20
         Top             =   3360
         Width           =   1335
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
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
         Left            =   480
         TabIndex        =   19
         Top             =   2880
         Width           =   1335
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
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
         Left            =   480
         TabIndex        =   18
         Top             =   2280
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Product ID"
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
         Left            =   0
         TabIndex        =   17
         Top             =   1800
         Width           =   2055
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Date Released"
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
         Left            =   7680
         TabIndex        =   16
         Top             =   1680
         Width           =   1815
      End
   End
   Begin VB.Frame Frame2 
      Height          =   2775
      Left            =   600
      TabIndex        =   3
      Top             =   4560
      Width           =   14055
      Begin MSComctlLib.ListView lvSales 
         Height          =   2415
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   13815
         _ExtentX        =   24368
         _ExtentY        =   4260
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
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Product ID"
            Object.Width           =   2647
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Product Name"
            Object.Width           =   5821
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Quantity"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Amount"
            Object.Width           =   2118
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Date Released"
            Object.Width           =   2293
         EndProperty
      End
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "New"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   12360
      Picture         =   "frmSales.frx":1AE4A
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   8040
      Width           =   2175
   End
   Begin VB.CommandButton cmdStockout 
      Caption         =   "Purchase"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   1080
      Picture         =   "frmSales.frx":317FC
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   8040
      Width           =   2295
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   9480
      Picture         =   "frmSales.frx":481AE
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   8040
      Width           =   2295
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "Transaction No."
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
      Left            =   1440
      TabIndex        =   39
      Top             =   1320
      Width           =   1935
   End
   Begin VB.Label lblChange 
      BackColor       =   &H80000007&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   11880
      TabIndex        =   36
      Top             =   7440
      Width           =   2295
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "Change:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10800
      TabIndex        =   35
      Top             =   7560
      Width           =   1575
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Cash:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7680
      TabIndex        =   33
      Top             =   7560
      Width           =   1215
   End
   Begin VB.Label lblAmount 
      BackColor       =   &H80000007&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   5160
      TabIndex        =   31
      Top             =   7440
      Width           =   2295
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Amount:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      TabIndex        =   30
      Top             =   7560
      Width           =   1575
   End
   Begin VB.Label lblItem 
      BackColor       =   &H80000007&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   2160
      TabIndex        =   29
      Top             =   7440
      Width           =   1335
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Item(s):"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      TabIndex        =   28
      Top             =   7560
      Width           =   1575
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   " Product ID"
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
      Left            =   960
      TabIndex        =   26
      Top             =   840
      Width           =   2055
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "New Sales"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   480
      Left            =   810
      TabIndex        =   24
      Top             =   0
      Width           =   2055
   End
End
Attribute VB_Name = "frmSales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbProdID_Click()
Dim tmpID As String
Dim siPadd, siName As String
Dim siQty, siAmount, siPrice As Double

siName = cmbProdName.Text
siQty = Val(txtQuantity.Text)
siPrice = Val(txtPrice.Text)
siPadd = cmbProdID.Text
tmpID = cmbProdID.Text

rs.Open "select*from tblProduct where Product_ID='" & tmpID & "'", cn, 3, 3
If rs.RecordCount > 0 Then
    cmbProdName.Text = rs!Product_Name
    txtSupplier.Text = rs!Supplier
    txtCategory.Text = rs!Category
    txtPrice.Text = rs!Unit_Price
    txtStocks.Text = rs!Unit_In_Stock
    

End If
Set rs = Nothing
cmdAdd.Enabled = True
End Sub

Private Sub cmbProdName_Click()
Dim tmpNme As String
Dim siPadd, siName As String
Dim siQty, siAmount, siPrice As Double

siName = cmbProdName.Text
siQty = Val(txtQuantity.Text)
siPrice = Val(txtPrice.Text)
siPadd = cmbProdID.Text
tmpID = cmbProdID.Text

tmpNme = cmbProdName.Text
rs.Open "select*from tblProduct where Product_Name='" & tmpNme & "'", cn, 3, 3
If rs.RecordCount > 0 Then
    cmbProdID.Text = rs!Product_ID
    txtSupplier.Text = rs!Supplier
    txtCategory.Text = rs!Category
    txtPrice.Text = rs!Unit_Price
    txtStocks.Text = rs!Unit_In_Stock
    
End If
Set rs = Nothing
cmdAdd.Enabled = True

End Sub

Private Sub cmdAdd_Click()

Dim siPadd, siName As String
Dim siQty, siAmount, siPrice As Double
siName = cmbProdName.Text
siQty = Val(txtQuantity.Text)
siPrice = Val(txtPrice.Text)
siAmount = siQty * siPrice
siPadd = cmbProdID.Text

If txtQuantity.Text = "" Then
    If MsgBox("Please Enter a Quantity", vbExclamation, "Inventory System") = vbOK Then
        txtQuantity.SetFocus
    End If
ElseIf cmbProdID.Text = "" Then
    If MsgBox("Please Enter a Product ID", vbExclamation, "Inventory System") = vbOK Then
        optProdID.SetFocus
    End If
ElseIf cmbProdName.Text = "" Then
    If MsgBox("Please Enter a Product Name", vbExclamation, "Inventory System") = vbOK Then
        optProdName.SetFocus
    End If
    
Else

With lvSales
    .ListItems.Add , , siPadd
    .ListItems(.ListItems.Count).ListSubItems.Add , , siName
    .ListItems(.ListItems.Count).ListSubItems.Add , , siQty
    .ListItems(.ListItems.Count).ListSubItems.Add , , siAmount
    .ListItems(.ListItems.Count).ListSubItems.Add , , DPDate_Released.Value
End With
cmbProdID.Text = ""
cmbProdName.Text = ""
txtSupplier.Text = ""
txtCategory.Text = ""
txtStocks.Text = ""
txtPrice.Text = ""
txtQuantity.Text = ""
cmdTot.Enabled = True
End If
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdNew_Click()
i_Clear.cLearMe Me
i_Enable.Enable_Txt Me
cmdAdd.Enabled = True

End Sub

Private Sub cmdPayment_Click()
Dim a As Integer
Dim subtotal As Currency
If txtCash.Text = "" Then
        If MsgBox("Please Input Cash", vbExclamation, "Inventory System") = vbOK Then
        txtCash.SetFocus
        End If

Else

For a = 1 To lvSales.ListItems.Count
    'the column that stores the Amount
    subtotal = subtotal + Val(lvSales.ListItems(a).SubItems(3))
    
Next a

lblAmount.Caption = Format(subtotal, "currency")
'computes for the change of the customer
If (txtCash.Text >= subtotal) Then
lblChange.Caption = Val(txtCash.Text) - subtotal
lblChange.Caption = Format(lblChange.Caption, "currency")
cmdStockout.Enabled = True
Else
MsgBox "Sorry Kulang ang Bayad mo!!", vbExclamation + vkokonly, "Inventory Systemt"
txtCash.SetFocus
End If

End If
End Sub

Private Sub cmdRemove_Click()
'This will remove Sales Items list.
Dim i As Integer
For i = lvSales.ListItems.Count To 1 Step -1
    If lvSales.ListItems(i).Checked = True Then
        lvSales.ListItems.Remove i
    End If
Next i
If lvSales.ListItems.Count = 0 Then cmdStockout.Enabled = False
End Sub

Private Sub cmdSearch_Click()
Dim pName As String
Dim tmpNme As String
Dim siPadd, siName As String
Dim siQty, siAmount, siPrice As Double

siName = cmbProdName.Text
siQty = Val(txtQuantity.Text)
siPrice = Val(txtPrice.Text)
siPadd = cmbProdID.Text
tmpID = cmbProdID.Text
If optProdName.Value = True Then
        pName = InputBox("Please enter Product Name to search")
        rs.Open "select*from tblProduct where Product_Name='" & pName & "'", cn, 3, 3
        If rs.RecordCount > 0 Then
            cmbProdName.Text = rs!Product_Name
            cmbProdID.Text = rs!Product_ID
            txtSupplier.Text = rs!Supplier
            txtCategory.Text = rs!Category
            txtPrice.Text = rs!Unit_Price
            txtStocks.Text = rs!Unit_In_Stock
        Else
        MsgBox "Record not Found, Please Enter a Valid Product Name", vbExclamation + vbOKOnly, "Inventory System"
        End If
        Set rs = Nothing
        cmdAdd.Enabled = True
ElseIf optProdID.Value = True Then
        pName = InputBox("Please enter Product ID to search")
        rs.Open "select*from tblProduct where Product_ID='" & pName & "'", cn, 3, 3
        If rs.RecordCount > 0 Then
            cmbProdName.Text = rs!Product_Name
            cmbProdID.Text = rs!Product_ID
            txtSupplier.Text = rs!Supplier
            txtCategory.Text = rs!Category
            txtPrice.Text = rs!Unit_Price
            txtStocks.Text = rs!Unit_In_Stock
        Else
        MsgBox "Record not Found, Please Enter a Valid Product ID", vbExclamation + vbOKOnly, "Inventory System"
        End If
        Set rs = Nothing
        cmdAdd.Enabled = True

Else
MsgBox "Please Select A Button To Search For A Product!!", vbExclamation + vbOKOnly, "Inventory System"
End If
Call Srch_Rec(Me.Name, pName)
End Sub



Private Sub cmdStockout_Click()

ctr = 0
Dim soTmp, soTmpSC, soTmpQty, soTmpAmnt, soTmpDRls As String
soTmp = lblSO.Caption
soTmpDRls = DPDate_Released.Value

'Checks for insufficient stocks
For i = 1 To lvSales.ListItems.Count
    soTmpSC = lvSales.ListItems(i).Text
    soTmpQty = lvSales.ListItems(i).ListSubItems(2).Text
    soTmpAmnt = lvSales.ListItems(i).ListSubItems(3).Text
    
    'Outing the stock from Product table
    rs.Open "Select*from tblProduct where Product_ID='" & soTmpSC & "'", cn, 3, 3
    If Val(rs!Unit_In_Stock) < CInt(soTmpQty) Then
        ctr = ctr + 1
    End If
    Set rs = Nothing
Next

If ctr = 0 Then
    For i = 1 To lvSales.ListItems.Count
        soTmpSC = lvSales.ListItems(i).Text
        soTmpQty = lvSales.ListItems(i).ListSubItems(2).Text
        soTmpAmnt = lvSales.ListItems(i).ListSubItems(3).Text
        
        rs.Open "Select*from tblProduct where Product_ID='" & soTmpSC & "'", cn, 3, 3
        rs!Unit_In_Stock = Val(rs!Unit_In_Stock) - soTmpQty
        rs.Update
        Set rs = Nothing
        
        cn.Execute "Insert Into tblSales(Transaction_No,Product_ID,Quantity,Amount,Date_Release)" & _
        "Values('" & soTmp & "','" & soTmpSC & "','" & soTmpQty & "','" & soTmpAmnt & "','" & soTmpDRls & "')"
    Next
    lvSales.ListItems.Clear
    cmdStockout.Enabled = False
    i_Clear.cLearMe Me
    i_Disable.Disable_Txt Me
    cmdAdd.Enabled = False
    soTmpSC = ""
    soTmpQty = ""
    soTmpAmnt = ""
    soTmpDRls = ""
    ctr = 0
    MsgBox "Stock out transaction No " & soTmp & " has been done.", vbInformation, "Inventory system"
    Call soLoadNum
Else
    MsgBox "Stock out transaction No " & soTmp & " has not been done due to" & vbCrLf & "some of the of the stocks are insufficient. Please check your" & vbCrLf & "quantity to be outed.", vbExclamation, "Inventory system"
    txtQuantity.SetFocus
    

End If

lblAmount.Caption = ""
lblChange.Caption = ""


End Sub


Private Sub cmdTot_Click()
Dim a, b As Integer
Dim subtotal As Currency
Dim totItem As Double

For a = 1 To lvSales.ListItems.Count
    'the column that stores your Amount
    subtotal = subtotal + Val(lvSales.ListItems(a).SubItems(3))
    
Next a

    lblAmount.Caption = Format(subtotal, "currency")
    lblItem.Caption = lvSales.ListItems.Count
    
 
 cmdPayment.Enabled = True
 
End Sub

Private Sub DataList1_Click()
rs.Open "Select*from tblProduct", cn, 3, 3

End Sub

Private Sub Form_Load()
Label9.Visible = False
DPDate_Released.Visible = False
cmbProdID.Locked = True
cmbProdName.Locked = True
txtSupplier.Locked = True
txtCategory.Locked = True
txtPrice.Locked = True
txtStocks.Locked = True


DPDate_Released.Value = Now
Call soLoadNum

'Loads up the product ID
rs.Open "Select*from tblProduct", cn, 3, 3
cmbProdID.Clear
If rs.RecordCount > 0 Then
    Do Until rs.EOF
        cmbProdID.AddItem rs!Product_ID
        rs.MoveNext
    Loop
End If
Set rs = Nothing

'Loads up the product names
rs.Open "Select*from tblProduct", cn, 3, 3
cmbProdName.Clear
If rs.RecordCount > 0 Then
    Do Until rs.EOF
        cmbProdName.AddItem rs!Product_Name
        rs.MoveNext
    Loop
End If
Set rs = Nothing

End Sub


'auto numberer
Private Function soLoadNum()
rs.Open "select * from tblSales Order By Transaction_No DESC", cn, 3, 2
If rs.RecordCount = 0 Then
    lblSO.Caption = "TN-0000"
Else
    lblSO.Caption = "TN-" & Format(Right(rs!Transaction_NO, 4) + 1, "0000")
End If
rs.Close
End Function







Private Sub lstProduct_Click()
Dim pNam, pID, pStock As String
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
