

Private Sub btn_delete_Click()
If Selected_List2 = 0 Then

    MsgBox "No row is selected.", vbOKOnly + vbInformation, "Delete"
    Exit Sub

    End If

    Dim i As VbMsgBoxResult

    i = MsgBox("Do you want to delete the selected record?", vbYesNo + vbQuestion, "Confirmation")

    If i = vbNo Then Exit Sub


    ThisWorkbook.Sheets("User").Rows(Selected_List2 + 1).Delete

    Call ResetPasswordForm

    MsgBox "Selected record has been deleted.", vbOKOnly + vbInformation, "Deleted"
End Sub

Private Sub btn_update_Click()
Dim Wk As String
    Dim BM As String
    Dim Ab As String

    Wk = frm_password.txt_username.Value
    BM = frm_password.txt_password.Value

    If Wk = "" Then
    MsgBox "Please double click on username or password !", vbOKOnly + vbCritical, "Error"
    Exit Sub
    ElseIf BM = "" Then
    MsgBox "Please double click on username or password!", vbOKOnly + vbCritical, "Error"
    Exit Sub
    Else
    Call UpdatePasswordForm

    End If

ActiveWorkbook.Save

MsgBox "Updated Successully!"

Call ResetPasswordForm
End Sub



Private Sub lst_usernampass_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
If Selected_List2 = 0 Then

    MsgBox "Kindly double click to edit.", vbOKOnly + vbInformation, "Double Click"
    Exit Sub

    End If

    'Code to update the value to respective controls

    Me.txtrownumber.Value = Selected_List2

    Me.txt_username = Me.lst_usernampass.List(Me.lst_usernampass.ListIndex, 0)

    Me.txt_password = Me.lst_usernampass.List(Me.lst_usernampass.ListIndex, 3)


    MsgBox "Please make the required changes and click on 'Add Entry' button to update."

End Sub


Private Sub txt_searchuserPass_Change()
Application.ScreenUpdating = False

Dim Database(1 To 1000000, 1 To 7)
Dim My_range As Integer
Dim Colum As Byte
On Error Resume Next

Me.lst_usernampass.RowSource = ""


For i = 1 To Sheet4.Range("A1000000").End(xlUp).Row

Sheet3.Range("A1").AutoFilter Field:=1, Criteria1:=LCase(Me.txt_searchuserPass.Value) & "*"

If LCase(Sheet4.Cells(i, 1)) Like LCase(Me.txt_searchuserPass.Value) & "*" = True Then
My_range = My_range + 1

For Colum = 1 To 7
Database(My_range, Colum) = Sheet4.Cells(i, Colum)

Next Colum
End If
Next i

Me.lst_usernampass.List = Database
Me.lst_usernampass.SetFocus
'Me.lst_database.ColumnHeads = False
'If Me.txt_searchrecord.Value = "" Then
'Me.lst_database.RowSource = "Database"
'lst_database.SetFocus
'Me.lst_database.ColumnHeads = False
'End If
Application.ScreenUpdating = True
End Sub

Private Sub UserForm_Activate()
Call msgpopup
End Sub

Private Sub UserForm_Initialize()


Call ResetPasswordForm

End Sub
