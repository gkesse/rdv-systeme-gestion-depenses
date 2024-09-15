Attribute VB_Name = "PassNpermissionsMODULE"
Sub ResetPasswordForm()
Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    
    With frm_password

    .txt_password.Value = ""

    .txt_username.Value = ""
    
    .txt_searchuserPass.Value = ""
    
    .txtrownumber.Value = ""
    
    End With

    Dim irow As Long
    irow = [counta(User!A:A)] 'identifying the last row
    
    With frm_password
    
    .lst_usernampass.ColumnCount = 7
    .lst_usernampass.ColumnHeads = True
    .lst_usernampass.ColumnWidths = "55,70,65,50,55,55,60"

    End With
    

    If irow > 1 Then
        frm_password.lst_usernampass.RowSource = "User!A2:G" & irow
    Else

    frm_password.lst_usernampass.RowSource = "User!A2:G"

    End If
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

End Sub

Sub UpdatePasswordForm()
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
    
    Dim Sh As Worksheet
    Dim irow As Long
    
    Set Sh = ThisWorkbook.Sheets("User")
    
    If frm_password.txtrownumber.Value = "" Then
    
    irow = [counta(User!A:A)] + 1
    
    Else: irow = frm_password.txtrownumber.Value + 1
    
    End If
    
    
    With Sh
    
   .Cells(irow, 1) = frm_password.txt_username.Value
    
    .Cells(irow, 3) = frm_password.txt_password.Value
    
    
    End With
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

End Sub

Function Selected_List2() As Long

Dim i As Long
Selected_List2 = 0

For i = 0 To frm_password.lst_usernampass.ListCount - 1

    If frm_password.lst_usernampass.Selected(i) = True Then

        Selected_List2 = i + 1
        Exit For

        End If
Next i

End Function
Sub ResetPermissionForm()
Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    
    With frm_permission

    .cbo_Analysis.Value = ""

    .cbo_dashboard.Value = ""
    
    .cbo_sysadmin.Value = ""
    
    .txtrownumber2.Value = ""
    
    End With

    Dim irow As Long
    irow = [counta(User!A:A)] 'identifying the last row
    
    With frm_permission
    
    .lst_perm.ColumnCount = 7
    .lst_perm.ColumnHeads = True
    .lst_perm.ColumnWidths = "55,70,65,50,55,55,60"

    End With
    

    If irow > 1 Then
        frm_permission.lst_perm.RowSource = "User!A2:G" & irow
    Else

    frm_permission.lst_perm.RowSource = "User!A2:G"

    End If
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

End Sub

Sub UpdatePermissionForm()
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
    
    Dim Sh As Worksheet
    Dim irow As Long
    
    Set Sh = ThisWorkbook.Sheets("User")
    
    If frm_permission.txtrownumber2.Value = "" Then
    
    irow = [counta(User!A:A)] + 1
    
    Else: irow = frm_permission.txtrownumber2.Value + 1
    
    End If
    
    
    With Sh
    
   .Cells(irow, 5) = frm_permission.cbo_Analysis.Value
    
    .Cells(irow, 6) = frm_permission.cbo_dashboard.Value
    
    .Cells(irow, 7) = frm_permission.cbo_sysadmin.Value
    
    End With
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

End Sub
Function Selected_List3() As Long

Dim i As Long
Selected_List3 = 0

For i = 0 To frm_permission.lst_perm.ListCount - 1

    If frm_permission.lst_perm.Selected(i) = True Then

        Selected_List3 = i + 1
        Exit For

        End If
Next i

End Function



Sub msgpopup()

CreateObject("Wscript.shell").Popup "Kindly double click on the database to make changes", 3, "Alert"

End Sub
