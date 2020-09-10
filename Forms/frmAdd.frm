VERSION 5.00
Begin VB.Form frmAdd 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add Form"
   ClientHeight    =   2595
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5160
   Icon            =   "frmAdd.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2595
   ScaleWidth      =   5160
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox lblProd_Name 
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
      ForeColor       =   &H80000002&
      Height          =   435
      Left            =   2040
      TabIndex        =   7
      Text            =   "Text2"
      Top             =   600
      Width           =   2895
   End
   Begin VB.TextBox lblProd_ID 
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
      ForeColor       =   &H80000002&
      Height          =   435
      Left            =   2040
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   120
      Width           =   2895
   End
   Begin VB.CommandButton cmdCancel 
      Height          =   780
      Left            =   3840
      Picture         =   "frmAdd.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1680
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Height          =   780
      Left            =   2160
      Picture         =   "frmAdd.frx":6704
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1680
      Width           =   975
   End
   Begin VB.TextBox txtOrderQty 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2040
      TabIndex        =   0
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Product Code:"
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
      Left            =   120
      TabIndex        =   5
      Top             =   240
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Product Name:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   2175
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Order Quantity"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   2175
   End
End
Attribute VB_Name = "frmAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()
Dim siStock As Double
'original
rs.Open "Select*from tblProduct where Product_ID='" & scAd & "'", cn, 3, 3
If rs.RecordCount > 0 Then
    qtyAd = txtOrderQty.Text
    With frmPOrder.lvOrderedList
        .ListItems.Add , , scAd
        .ListItems(.ListItems.Count).ListSubItems.Add , , itmAd
        .ListItems(.ListItems.Count).ListSubItems.Add , , qtyAd
        .ListItems(.ListItems.Count).ListSubItems.Add , , rs!Supplier
        .ListItems(.ListItems.Count).ListSubItems.Add , , rs!Category
        .ListItems(.ListItems.Count).ListSubItems.Add , , rs!Unit_Price
        .ListItems(.ListItems.Count).ListSubItems.Add , , rs!Unit_In_Stock
    End With
    frmPOrder.lblstat.Caption = frmPOrder.lvOrderedList.ListItems.Count
    
End If

Set rs = Nothing
scAd = ""
itmAd = ""
qtyAd = ""
Unload Me
frmPOrder.cmdAdd.Enabled = True

End Sub


