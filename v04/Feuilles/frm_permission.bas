Private Sub btn_ok_Click()
    Dim Wk As String
    Dim BM As String
    Dim Ab As String
    Dim Dc As String


    Wk = frm_permission.cbo_Analysis.Value
    BM = frm_permission.cbo_dashboard.Value
    Dc = frm_permission.cbo_sysadmin.Value

    If Wk = "" Then
        MsgBox "Please double click on permissions !", vbOKOnly + vbCritical, "Error"
        Exit Sub
        ElseIf BM = "" Then
        MsgBox "Please double click on permissions !", vbOKOnly + vbCritical, "Error"
        Exit Sub
        ElseIf Dc = "" Then
        MsgBox "Please double click on permissions !", vbOKOnly + vbCritical, "Error"
        Exit Sub
    Else

        Call UpdatePermissionForm

    End If

ActiveWorkbook.Save

MsgBox "Updated Successully!"

Call ResetPermissionForm
End Sub

Private Sub lst_perm_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

If Selected_List3 = 0 Then

    MsgBox "Kindly double click to edit.", vbOKOnly + vbInformation, "Double Click"
    Exit Sub

End If

    'Code to update the value to respective controls

    Me.txtrownumber2.Value = Selected_List3

    Me.cbo_Analysis = Me.lst_perm.List(Me.lst_perm.ListIndex, 4)

    Me.cbo_dashboard = Me.lst_perm.List(Me.lst_perm.ListIndex, 5)

    Me.cbo_sysadmin = Me.lst_perm.List(Me.lst_perm.ListIndex, 6)

    MsgBox "Please make the required changes and click on button to update."
End Sub

Private Sub txt_searchuserPerm_Change()
Application.ScreenUpdating = False

Dim Database(1 To 1000000, 1 To 7)
Dim My_range As Integer
Dim Colum As Byte
On Error Resume Next

Me.lst_perm.RowSource = ""


For i = 1 To Sheet4.Range("A1000000").End(xlUp).Row

    Sheet3.Range("A1").AutoFilter Field:=1, Criteria1:=LCase(Me.txt_searchuserPerm.Value) & "*"

    If LCase(Sheet4.Cells(i, 1)) Like LCase(Me.txt_searchuserPerm.Value) & "*" = True Then
        My_range = My_range + 1

        For Colum = 1 To 7
            Database(My_range, Colum) = Sheet4.Cells(i, Colum)

        Next Colum
    End If
Next i

Me.lst_perm.List = Database
Me.lst_perm.SetFocus

Application.ScreenUpdating = True
End Sub

Private Sub UserForm_Activate()
Call msgpopup
End Sub

Private Sub UserForm_Initialize()

Call ResetPermissionForm

With Me.cbo_Analysis

.AddItem "TRUE"
.AddItem "FALSE"

End With

With Me.cbo_dashboard

.AddItem "TRUE"
.AddItem "FALSE"

End With

With Me.cbo_sysadmin

.AddItem "TRUE"
.AddItem "FALSE"

End With


End Sub
