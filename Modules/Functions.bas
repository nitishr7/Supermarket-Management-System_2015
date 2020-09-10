Attribute VB_Name = "Functions"

'Function to search record
Function Srch_Rec(frm_Name, Fld_Param)
Select Case frm_Name
    Case "frmProducts":
        rs.Open "Select* from tblProduct", cn, 3, 3
            Do While Not rs.EOF
                If Fld_Param = rs!Product_ID Then
                    With frmProducts
                        .txtProduct_ID.Text = rs!Product_ID
                        .txtProduct_Name = rs!Product_Name
                        .cboSupplier.Text = rs!Supplier
                        .cboCategory.Text = rs!Category
                        .txtUnit_Price.Text = rs!Unit_Price
                        .txtINStock.Text = rs!Unit_In_Stock
                        ctr = 0
                        
                        
                    Exit Do
                    End With
                End If
                rs.MoveNext
                ctr = ctr + 1
            Loop
            If ctr > 0 Then
                MsgBox "No record found", vbInformation
            End If
        Set rs = Nothing
End Select
End Function


'Function to save and update records
Function SQL_Execute(ctrl_Flag, prod_ID)
    With frmProducts
        Prod_Code = .txtProduct_ID.Text
        Prod_Desc = .txtProduct_Name.Text
        Sp = .cboSupplier.Text
        Cat = .cboCategory.Text
        U_Price = .txtUnit_Price.Text
        U_N_Stock = .txtINStock.Text
    End With

Select Case ctrl_Flag
    Case False:
        rs.Open "Insert into tblProduct(Product_ID,Product_Name,Supplier,Category,Unit_Price,Unit_In_Stock)" & _
        "values('" & Prod_Code & "','" & Prod_Desc & "','" & Sp & "','" & Cat & "','" & U_Price & "','" & U_N_Stock & "')", cn, adOpenKeyset, adLockPessimistic
        Set rs = Nothing
        MsgBox "The product named " & Prod_Desc & " Has been save successfully.", vbInformation
    Case True:
        rs.Open "Update tblProduct set Product_Name='" & Prod_Desc & "',Supplier='" & Sp & "',Category='" & Cat & "',Unit_Price='" & U_Price & "',Unit_In_Stock='" & U_N_Stock & "'" & _
        "Where Product_ID='" & prod_ID & "'", cn, adOpenKeyset, adLockPessimistic
        Set rs = Nothing
        MsgBox "The product with Code no.  " & prod_ID & " Has been updated successfully.", vbInformation
End Select
    Call Reload_List
End Function


Function Reload_List()
'Reload Listbox
frmProducts.lstProduct.Clear
rs.Open "Select*from tblProduct", cn, 3, 3
If rs.RecordCount > 0 Then
    Do While Not rs.EOF
        frmProducts.lstProduct.AddItem rs!Product_Name
        rs.MoveNext
    Loop
End If
Set rs = Nothing
End Function


'############### FIND SUPPLIER PROFILE ##################
Function sup_Find()
ctr = 0
rs.Open "Select*from tblSupplier", cn, 3, 3
If rs.RecordCount > 0 Then
    Do Until rs.EOF
        If iFind = rs!supplier_id Then
            With frmSupplier
                .txtSupID.Text = rs!supplier_id
                .txtSupName.Text = rs!supplier_name
                .txtAddress.Text = rs!Address
                .txtTelephone.Text = rs!Telephone
                ctr = 0
                Set rs = Nothing
                Exit Do
            End With
        Else
            rs.MoveNext
            ctr = ctr + 1
        End If
    Loop
    If ctr > 0 Then
        Set rs = Nothing
        MsgBox "The supplier with Supplier ID " & iFind & " is not found.", vbExclamation, "Inventory System"
    End If
End If
End Function


'############### FIND USER PROFILE  #######################
Function usr_Find()
ctr = 0
rs.Open "Select*from tblEmployee", cn, 3, 3
If rs.RecordCount > 0 Then
    Do Until rs.EOF
        If iFind = rs!Employee_ID Then
            With frmUser
                .txtUsrID.Text = rs!Employee_ID
                .txtFulName.Text = rs!Employee_Name
                .txtDesignation.Text = rs!designation
                .txtDepartment.Text = rs!department
                .txtPwd.Text = rs!Password
                ctr = 0
                Set rs = Nothing
                Exit Do
            End With
        Else
            rs.MoveNext
            ctr = ctr + 1
        End If
    Loop
    If ctr > 0 Then
        Set rs = Nothing
        MsgBox "The user ID " & iFind & " is not found.", vbExclamation, "Inventory System"
    End If
End If


End Function








