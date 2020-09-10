VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "Supermarket Management System"
   ClientHeight    =   8460
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   14040
   Icon            =   "Main.frx":0000
   LinkTopic       =   "MDIForm1"
   MouseIcon       =   "Main.frx":6852
   Picture         =   "Main.frx":D0A4
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Top             =   7965
      Width           =   14040
      _ExtentX        =   24765
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5292
            MinWidth        =   5292
            Picture         =   "Main.frx":7AEF7
            Text            =   "User"
            TextSave        =   "User"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5292
            MinWidth        =   5292
            Picture         =   "Main.frx":8FE52
            Text            =   "Ready"
            TextSave        =   "Ready"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Enabled         =   0   'False
            Object.Width           =   5292
            MinWidth        =   5292
            Picture         =   "Main.frx":95FDC
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Object.Width           =   5292
            MinWidth        =   5292
            Picture         =   "Main.frx":9C166
            TextSave        =   "2:17 PM"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Object.Width           =   5292
            MinWidth        =   5292
            Picture         =   "Main.frx":A22F0
            TextSave        =   "10/11/2016"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   840
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   14040
      _ExtentX        =   24765
      _ExtentY        =   1482
      ButtonWidth     =   1826
      ButtonHeight    =   1429
      Appearance      =   1
      Style           =   1
      ImageList       =   "i32x32"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   16
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Transaction"
            ImageIndex      =   30
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Product"
            ImageIndex      =   32
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Supplier"
            ImageIndex      =   31
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "User"
            ImageIndex      =   28
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   5
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
               BeginProperty ButtonMenu5 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Cashier"
            ImageIndex      =   27
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Report"
            ImageIndex      =   24
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Lock"
            ImageIndex      =   23
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Exit"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList i32x32 
      Left            =   720
      Top             =   1200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   32
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":A847A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":A9154
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":A96EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":A9C88
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":AA222
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":AAEFC
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":ABBD6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":AC8B0
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":AD58A
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":AE264
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":AEF3E
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":AFC18
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":B08F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":B15CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":B22A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":B2F80
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":C73C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":11E56C
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":1205FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":127B00
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":149D60
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":15EED2
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":174044
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":1891B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":19E328
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":1B349A
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":1B9624
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":1BF7AE
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":1D5952
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":1EC9EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":1F21DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":1F79D0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu file 
      Caption         =   "File"
      Begin VB.Menu lin 
         Caption         =   "Log In"
      End
      Begin VB.Menu lot 
         Caption         =   "Log Out"
         Enabled         =   0   'False
      End
      Begin VB.Menu bar1 
         Caption         =   "-"
      End
      Begin VB.Menu loc 
         Caption         =   "Lock"
         Enabled         =   0   'False
      End
      Begin VB.Menu bar7 
         Caption         =   "-"
      End
      Begin VB.Menu ext 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu trans 
      Caption         =   "Transaction"
      Enabled         =   0   'False
      Begin VB.Menu sin 
         Caption         =   "Stock In"
      End
      Begin VB.Menu sot 
         Caption         =   "Stock Out"
      End
      Begin VB.Menu bar2 
         Caption         =   "-"
      End
      Begin VB.Menu po 
         Caption         =   "Purchase Order"
      End
   End
   Begin VB.Menu maint 
      Caption         =   "Maintenance"
      Enabled         =   0   'False
      Begin VB.Menu prod 
         Caption         =   "Product"
      End
      Begin VB.Menu sup 
         Caption         =   "Supplier"
      End
      Begin VB.Menu bar3 
         Caption         =   "-"
      End
      Begin VB.Menu user 
         Caption         =   "User"
      End
   End
   Begin VB.Menu rep 
      Caption         =   "Report"
      Enabled         =   0   'False
      Begin VB.Menu si 
         Caption         =   "Product Stock in"
      End
      Begin VB.Menu so 
         Caption         =   "Product Stock out"
      End
      Begin VB.Menu bar8 
         Caption         =   "-"
      End
      Begin VB.Menu aps 
         Caption         =   "All Product by Supplier"
      End
      Begin VB.Menu apc 
         Caption         =   "All Product by Category"
      End
      Begin VB.Menu bar4 
         Caption         =   "-"
      End
      Begin VB.Menu ap 
         Caption         =   "All Product"
      End
   End
   Begin VB.Menu tol 
      Caption         =   "Tools"
      Begin VB.Menu clk 
         Caption         =   "Clock"
      End
      Begin VB.Menu calc 
         Caption         =   "Calculator"
      End
      Begin VB.Menu pad 
         Caption         =   "Notepad"
      End
   End
   Begin VB.Menu about 
      Caption         =   "About"
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub about_Click()
frmAbout.Show
End Sub

Private Sub ap_Click()
On Error Resume Next

Set rs = New ADODB.Recordset
rs.Open "SELECT * From tblProduct", cn
Set DataReport1.DataSource = rs.DataSource
For Each obj In DataReport1.Sections("Section1").Controls
    If TypeOf obj Is RptTextBox Then
        obj.DataMember = rs.DataMember
    End If
Next
DataReport1.Sections("Section1").Controls("Text1").DataField = "Product_ID"
DataReport1.Sections("Section1").Controls("Text2").DataField = "Product_Name"
DataReport1.Sections("Section1").Controls("Text3").DataField = "Supplier"
DataReport1.Sections("Section1").Controls("Text4").DataField = "Category"
DataReport1.Sections("Section1").Controls("Text5").DataField = "Unit_Price"
DataReport1.Sections("Section1").Controls("Text6").DataField = "Unit_In_Stock"
DataReport1.Refresh
DataReport1.Show
Set rs = Nothing
End Sub



Private Sub apc_Click()
On Error Resume Next
Dim RPT$, RPT2$
RPT = InputBox("Enter Product Category.")

Set rs = New ADODB.Recordset
rs.Open "SELECT * From tblProduct where Category='" & RPT & "'", cn
RPT2 = rs!Category
Set DataReport3.DataSource = rs.DataSource

For Each obj In DataReport3.Sections("Section1").Controls
    If TypeOf obj Is RptTextBox Then
        obj.DataMember = rs.DataMember
    End If
Next
DataReport3.Sections("Section1").Controls("Text1").DataField = "Product_ID"
DataReport3.Sections("Section1").Controls("Text2").DataField = "Product_Name"
DataReport3.Sections("Section2").Controls("Label1").Caption = RPT2
DataReport3.Sections("Section1").Controls("Text3").DataField = "Supplier"
DataReport3.Sections("Section1").Controls("Text5").DataField = "Unit_Price"
DataReport3.Sections("Section1").Controls("Text6").DataField = "Unit_In_Stock"
DataReport3.Refresh
DataReport3.Show
Set rs = Nothing
End Sub

Private Sub aps_Click()
On Error Resume Next
Dim RPT$, RPT2$
RPT = InputBox("Enter product supplier name.")

Set rs = New ADODB.Recordset
rs.Open "SELECT * From tblProduct where Supplier='" & RPT & "'", cn
RPT2 = rs!Supplier
Set DataReport2.DataSource = rs.DataSource

For Each obj In DataReport2.Sections("Section1").Controls
    If TypeOf obj Is RptTextBox Then
        obj.DataMember = rs.DataMember
    End If
Next
DataReport2.Sections("Section1").Controls("Text1").DataField = "Product_ID"
DataReport2.Sections("Section1").Controls("Text2").DataField = "Product_Name"
DataReport2.Sections("Section2").Controls("Label1").Caption = RPT2
DataReport2.Sections("Section1").Controls("Text4").DataField = "Category"
DataReport2.Sections("Section1").Controls("Text5").DataField = "Unit_Price"
DataReport2.Sections("Section1").Controls("Text6").DataField = "Unit_In_Stock"
DataReport2.Refresh
DataReport2.Show
Set rs = Nothing
End Sub

Private Sub calc_Click()
Shell "calc.exe", vbNormalFocus
End Sub

Private Sub clk_Click()
Shell "rundll32.exe shell32.dll,Control_RunDLL timedate.cpl", vbNormalFocus

End Sub

Private Sub ext_Click()
End
End Sub

Private Sub lin_Click()
frmLogin.Show
End Sub

Private Sub loc_Click()
frmlock.Show
End Sub

Private Sub lot_Click()
lin.Enabled = True
trans.Enabled = False
maint.Enabled = False
rep.Enabled = False
lot.Enabled = False
Toolbar1.Buttons(1).Enabled = False
Toolbar1.Buttons(3).Enabled = False
Toolbar1.Buttons(5).Enabled = False
Toolbar1.Buttons(7).Enabled = False
Toolbar1.Buttons(9).Enabled = False
Toolbar1.Buttons(13).Enabled = False
End Sub

Private Sub MDIForm_Load()
dBase = App.Path & "\Inventory.mdb"
cn.Open "Driver={Microsoft Access Driver (*.mdb)};dbq=" & dBase
Toolbar1.Buttons(1).Enabled = False
Toolbar1.Buttons(3).Enabled = False
Toolbar1.Buttons(5).Enabled = False
Toolbar1.Buttons(7).Enabled = False
Toolbar1.Buttons(9).Enabled = False
Toolbar1.Buttons(13).Enabled = False

frmLogin.Show vbModal
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
End
End Sub

Private Sub pad_Click()
Shell "notepad.exe", vbNormalFocus
End Sub



Private Sub Picture1_Resize()
 'Picture1.Picture = Image1(1).Picture
    Picture1.ScaleMode = 3
    Picture1.AutoRedraw = True
    Picture1.PaintPicture Picture1.Picture, _
        0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, _
        0, 0, _
        Picture1.Picture.Width / 26.46, _
        Picture1.Picture.Height / 26.46
    Picture1.Picture = Picture1.Image
End Sub



Private Sub prod_Click()
frmProducts.Show vbModal
End Sub

Private Sub si_Click()
On Error Resume Next

Set rs = New ADODB.Recordset
rs.Open "SELECT * From tblStockIn", cn
Set DataReport4.DataSource = rs.DataSource
For Each obj In DataReport4.Sections("Section1").Controls
    If TypeOf obj Is RptTextBox Then
        obj.DataMember = rs.DataMember
    End If
Next
DataReport4.Sections("Section1").Controls("Text1").DataField = "SI_No"
DataReport4.Sections("Section1").Controls("Text2").DataField = "PO_No"
DataReport4.Sections("Section1").Controls("Text3").DataField = "Date_Recieved"
DataReport4.Sections("Section1").Controls("Text4").DataField = "Product_ID"
DataReport4.Sections("Section1").Controls("Text5").DataField = "Quantity"
DataReport4.Refresh
DataReport4.Show
Set rs = Nothing
End Sub

Private Sub sin_Click()
frmStockIn.Show vbModal
End Sub

Private Sub so_Click()
On Error Resume Next

Set rs = New ADODB.Recordset
rs.Open "SELECT * From tblStockout", cn
Set DataReport5.DataSource = rs.DataSource
For Each obj In DataReport5.Sections("Section1").Controls
    If TypeOf obj Is RptTextBox Then
        obj.DataMember = rs.DataMember
    End If
Next
DataReport5.Sections("Section1").Controls("Text1").DataField = "SO_No"
DataReport5.Sections("Section1").Controls("Text2").DataField = "Product_ID"
DataReport5.Sections("Section1").Controls("Text3").DataField = "Quantity"
DataReport5.Sections("Section1").Controls("Text4").DataField = "Amount"
DataReport5.Sections("Section1").Controls("Text5").DataField = "Date_Release"
DataReport5.Refresh
DataReport5.Show
Set rs = Nothing
End Sub

Private Sub sot_Click()
frmStockout.Show vbModal
End Sub

Private Sub sup_Click()
frmSupplier.Show vbModal
End Sub

Private Sub po_Click()
frmPOrder.Show vbModal
End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
Case 1:
PopupMenu trans, , Button.Left, (Button.Top + Button.Height)
Case 3:
frmProducts.Show vbModal
Case 5:
frmSupplier.Show vbModal
Case 7:
frmAdminAutorize.Show vbModal
Case 9:
frmSales.Show vbModal
Case 11: PopupMenu rep, , Button.Left, (Button.Top + Button.Height)
Case 13:
frmlock.Show
Case 15:
If MsgBox("Do you want to Exit??", vbQuestion + vbYesNo, "Inventory System") = vbYes Then
End
Else
MDIForm1.Show
End If
End Select
End Sub

Private Sub user_Click()
frmAdminAutorize.Show vbModal
End Sub
