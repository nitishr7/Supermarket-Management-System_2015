Attribute VB_Name = "Global"
Public cn As New adodb.Connection
Public rs As New adodb.Recordset
Public dBase As String

Public ctrl_Flag As Boolean
Public iTerminate As Boolean

Public i_Clear As New TextHandling
Public i_Enable As New TextHandling
Public i_Disable As New TextHandling
Public Field_Check As New TextHandling


Public obj As Control
Public prod_ID As String
Public i_Delete As String
Public iFind As String
Public ctr As Integer
Public i, j, k, l As Integer
Public poN As String


Public Prod_Code, Prod_Desc, Sp, Cat, U_Price, U_N_Stock As String
Public Sup_ID, Sup_Name, sAdd, sTel As String
Public scAd, itmAd, qtyAd As String
