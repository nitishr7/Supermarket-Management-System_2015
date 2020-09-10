VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmStockIn 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Stock In"
   ClientHeight    =   8625
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11820
   Icon            =   "frmStockIn.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8625
   ScaleWidth      =   11820
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   6855
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   11535
      Begin VB.ComboBox Combo1 
         BackColor       =   &H00FFFFFF&
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
         ForeColor       =   &H00FF0000&
         Height          =   390
         Left            =   3000
         TabIndex        =   7
         Top             =   840
         Width           =   3735
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         ForeColor       =   &H00FF0000&
         Height          =   480
         Left            =   3000
         TabIndex        =   6
         Top             =   1440
         Width           =   3735
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         ForeColor       =   &H00FF0000&
         Height          =   420
         Left            =   9240
         TabIndex        =   5
         Top             =   240
         Width           =   2055
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         ForeColor       =   &H00FF0000&
         Height          =   420
         Left            =   9240
         TabIndex        =   4
         Top             =   840
         Width           =   2055
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   4335
         Left            =   120
         TabIndex        =   3
         Top             =   2160
         Width           =   11175
         _ExtentX        =   19711
         _ExtentY        =   7646
         _Version        =   393216
         Rows            =   1
         Cols            =   5
         FixedCols       =   0
         BackColor       =   16777215
         ForeColor       =   0
         BackColorFixed  =   12632256
         ForeColorFixed  =   0
         BackColorSel    =   12640511
         ForeColorSel    =   16761087
         BackColorBkg    =   16777215
         GridColor       =   4210752
         GridColorFixed  =   49152
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblStockIn 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   435
         Left            =   3000
         TabIndex        =   13
         Top             =   240
         Width           =   3735
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Stock In Number"
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
         TabIndex        =   12
         Top             =   240
         Width           =   2775
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "PO No."
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
         TabIndex        =   11
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Supplier Code"
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
         TabIndex        =   10
         Top             =   1440
         Width           =   1815
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Date Ordered"
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
         Left            =   7080
         TabIndex        =   9
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Date Recieved"
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
         Left            =   7080
         TabIndex        =   8
         Top             =   840
         Width           =   1815
      End
   End
   Begin VB.CommandButton Command2 
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
      Height          =   855
      Left            =   8400
      Picture         =   "frmStockIn.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7560
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
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
      Height          =   855
      Left            =   10080
      Picture         =   "frmStockIn.frx":0B14
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7560
      Width           =   1575
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000009&
      Height          =   255
      Left            =   0
      Top             =   0
      Width           =   11775
   End
End
Attribute VB_Name = "frmStockIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public rs2 As New ADODB.Recordset


Private Sub Combo1_Click()
rs.Open "Select * from tblPO where PO_NO='" & Combo1.Text & "'", cn, 3, 2
MSFlexGrid1.Rows = 1
Do While Not rs.EOF
   Text1.Text = rs.Fields!Po_Supplier
   Text2.Text = rs.Fields![PO_Date_Order]
   Text3.Text = Date
   MSFlexGrid1.AddItem rs.Fields![Product_ID]
   MSFlexGrid1.TextMatrix(MSFlexGrid1.Rows - 1, 3) = rs.Fields![Quantity]
   rs2.Open "Select * from tblProduct where `Product_ID` ='" & rs.Fields![Product_ID] & "'", cn, 3, 3
   If rs2.RecordCount > 0 Then
   MSFlexGrid1.TextMatrix(MSFlexGrid1.Rows - 1, 1) = rs2.Fields![Product_Name]
   MSFlexGrid1.TextMatrix(MSFlexGrid1.Rows - 1, 2) = rs2.Fields![Unit_Price]
   
  End If
     rs2.Close
   rs.MoveNext
   Loop
   rs.Close
End Sub

Private Sub Command1_Click()
If Command1.Caption = "New" Then
  'enable all components
    Combo1.Enabled = True
    i_Enable.Enable_Txt Me
    MSFlexGrid1.Enabled = True
  'clear the components
    Combo1.Text = ""
    i_Clear.cLearMe Me
    MSFlexGrid1.Rows = 1
  'put all PONO to combo1 without duplicate
    rs.Open "select Distinct PO_NO from tblPO", cn, 3, 2
    Combo1.Clear
    Do While Not rs.EOF
        Combo1.AddItem rs.Fields!PO_NO
        rs.MoveNext
    Loop
    rs.Close
  'auto number of stock in number
    rs.Open "select * from tblStockIn Order By Si_No DESC", cn, 3, 2
    If rs.RecordCount = 0 Then
    lblStockIn.Caption = "SI-0000"
   Else
      lblStockIn.Caption = "SI-" & Format(Right(rs.Fields!SI_NO, 4) + 1, "0000")
   End If
   rs.Close
   Command1.Caption = "save"
   Combo1.SetFocus
  Else
  'put all stock in StockIn Table
   rs.Open "select * from tblStockin", cn, 3, 2
     For i = 1 To MSFlexGrid1.Rows - 1
         rs.AddNew
         rs.Fields!SI_NO = lblStockIn.Caption
         rs.Fields!PO_NO = Combo1.Text
         rs.Fields!Date_Recieved = Date
         rs.Fields![Product_ID] = MSFlexGrid1.TextMatrix(i, 0)
         rs.Fields!Quantity = MSFlexGrid1.TextMatrix(i, 4)
         rs.Update
        'add the Unit In Stock
         rs2.Open "select * from tblProduct where `Product_ID`='" & MSFlexGrid1.TextMatrix(i, 0) & "'", cn, 3, 2
         If rs2.RecordCount > 0 Then
           rs2.Fields![Unit_In_Stock] = Val(rs2.Fields![Unit_In_Stock]) + Val(MSFlexGrid1.TextMatrix(i, 4))
           rs2.Update
         End If
         rs2.Close
    Next i
   rs.Close
  
 'delete PO,that already stock in
   rs.Open "delete from tblPO where PO_NO='" & Combo1.Text & "'", cn, 3, 2
   Command1.Caption = "New"
 'clear and disable the components
    Combo1.Enabled = False
    i_Disable.Disable_Txt Me
    MSFlexGrid1.Enabled = False
    Combo1.Text = ""
    i_Clear.cLearMe Me
    MSFlexGrid1.Rows = 1
    lblStockIn.Caption = ""
End If
Command1.Enabled = False
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()

 MSFlexGrid1.TextMatrix(0, 0) = "Product ID"
   MSFlexGrid1.TextMatrix(0, 1) = "Name"
   MSFlexGrid1.TextMatrix(0, 2) = "Price"
   MSFlexGrid1.TextMatrix(0, 3) = "Order"
   MSFlexGrid1.TextMatrix(0, 4) = "Receive"
   
   
With MSFlexGrid1
    .ColWidth(0) = 1300
    .ColWidth(1) = 3300
    .ColWidth(2) = 1100
    .ColWidth(3) = 1100
    .ColWidth(4) = 1100
End With
   
'Put all supplier code into the combobox
If rs.State = 1 Then Set rs = Nothing
   rs.Open "Select Distinct `PO_No` from tblPO", cn, 3, 2
   Do While Not rs.EOF
        Combo1.AddItem rs.Fields!PO_NO
        rs.MoveNext
   Loop
   rs.Close
   
   
   
End Sub


Private Sub Label5_Click()

End Sub

Private Sub MSFlexGrid1_DblClick()
frmReceive.lblProductID.Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 0)
frmReceive.lblProductName.Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1)
'Load frmReceive
frmReceive.Show vbModal
End Sub

