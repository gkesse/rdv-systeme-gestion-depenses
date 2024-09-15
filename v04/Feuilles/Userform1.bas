Private Sub CMDCREATE_Click()
Dim DataUser As Object
Set DataUser = Sheet4.Range("A10000").End(xlUp)
If Me.TXTUSERNAME.Value = "" _
Or Me.TXTEMAIL.Value = "" _
Or Me.TXTPHONE.Value = "" _
Or Me.TXTNEWPASSWORD.Value = "" _
Or Me.TXTCONFIRMPASSWORD.Value = "" _
Or Me.TXTNEWPASSWORD.Value <> Me.TXTCONFIRMPASSWORD.Value Then
Call MsgBox("Password isn't consistent", vbInformation, "Please! Enter accurate password")
Else
DataUser.Offset(1, 0).Value = Me.TXTUSERNAME.Value
DataUser.Offset(1, 1).Value = Me.TXTEMAIL.Value
DataUser.Offset(1, 2).Value = Me.TXTPHONE.Value
DataUser.Offset(1, 3).Value = Me.TXTNEWPASSWORD.Value
If Me.ONOFF1.Caption = "ON" Then
DataUser.Offset(1, 4).Value = "TRUE"
End If
If Me.ONOFF1.Caption = "OFF" Then
DataUser.Offset(1, 4).Value = "FALSE"
End If

If Me.ONOFF2.Caption = "ON" Then
DataUser.Offset(1, 5).Value = "TRUE"
End If
If Me.ONOFF2.Caption = "OFF" Then
DataUser.Offset(1, 5).Value = "FALSE"
End If

If Me.ONOFF3.Caption = "ON" Then
DataUser.Offset(1, 6).Value = "TRUE"
End If
If Me.ONOFF3.Caption = "OFF" Then
DataUser.Offset(1, 6).Value = "FALSE"

End If

Call MsgBox("User registeration succesful", vbInformation, "New User")
Me.TXTUSERNAME.Value = ""
Me.TXTEMAIL.Value = ""
Me.TXTPHONE.Value = ""
Me.TXTNEWPASSWORD.Value = ""
Me.TXTCONFIRMPASSWORD.Value = ""
Call ButtonOff1
Call ButtonOff2
Call ButtonOff3
End If

ActiveWorkbook.Save

End Sub

Private Sub CMDLOGIN_Click()
'Sheet5.Range("A1").Select
On Error GoTo Incorrect
Set CariUser = Sheet4.Range("A2:A100").Find(What:=Me.TXTUSER.Value, LookIn:=xlValues)
Set DataUser = Sheet4.Range("A2:A100").Find(What:=Me.TXTUSER.Value, LookIn:=xlValues)

Me.TXTCHECKPASSWORD.Value = CariUser.Offset(0, 3).Value
Me.TXTCHECKUSERNAME.Value = CariUser.Offset(0, 0).Value

If Me.TXTUSER.Value = "" _
    Or Me.TXTPASSWORD.Value = "" _
    Or Me.TXTUSER.Value <> Me.TXTCHECKUSERNAME.Value _
    Or Me.TXTPASSWORD.Value <> Me.TXTCHECKPASSWORD.Value Then
    Me.Notifywrongpassword.Visible = True
'Call MsgBox("Kindly provide the correct Username or Password", vbInformation, "Login Error")
Else

    Application.Visible = True
    Call Permission
    Me.Notifywrongpassword.Visible = False
    'MsgBox "Access Granted", vbInformation, "User Account"
    Unload Me


    Sheet10.Range("A12").Value = Sheet10.Range("A11").Value & Me.TXTUSER.Value
    frm_expense.LBUSER.Caption = Sheet10.Range("A12").Value
    Sheet10.Range("A15").Value = Me.TXTUSER.Value

    Sheet5.Range("A1").Select
End If
'Application.Visible = True

Exit Sub
Incorrect:
Call MsgBox("Welcome " & Me.TXTUSER.Value, vbInformation, "Verified User")
Unload Me
End Sub
Private Sub Permission()

Set CariUser = Sheet4.Range("A2:A100").Find(What:=Me.TXTUSER.Value, LookIn:=xlValues)

If CariUser.Offset(0, 4).Value = False Then
Sheet5.Analysis.Enabled = False
Sheet7.Analysis2.Enabled = False
Sheet9.Analysis3.Enabled = False
Else
Sheet5.Analysis.Enabled = True
Sheet7.Analysis2.Enabled = True
Sheet9.Analysis3.Enabled = True
End If

If CariUser.Offset(0, 5).Value = False Then
Sheet5.Dashboard.Enabled = False
Sheet7.Dashboard2.Enabled = False
Sheet9.Dashboard3.Enabled = False
Else
Sheet5.Dashboard.Enabled = True
Sheet7.Dashboard2.Enabled = True
Sheet9.Dashboard3.Enabled = True
End If

If CariUser.Offset(0, 6).Value = False Then
Sheet5.SysAdmin.Enabled = False
Sheet7.SysAdmin2.Enabled = False
Sheet9.SysAdmin3.Enabled = False
Sheet5.Shapes("Rectangle 24").Visible = msoFalse
Sheet5.Shapes("Picture 21").Visible = msoFalse
Else
Sheet5.SysAdmin.Enabled = True
Sheet7.SysAdmin2.Enabled = True
Sheet9.SysAdmin3.Enabled = True
Sheet5.Shapes("Rectangle 24").Visible = msoTrue
Sheet5.Shapes("Picture 21").Visible = msoTrue
End If

End Sub


Private Sub CMDLOGIN1_Click()
Me.MultiPage1.Value = 0
End Sub

Private Sub CMDOK_Click()
If Me.TXTADMINPASSWORD.Value <> "admin1234" Then
Call MsgBox("Excuse Me, Password Admin Wrong, Kindly Contact Administrator", vbInformation, "Wrong Password")
Me.TXTADMINPASSWORD.Value = ""
Me.TXTADMINPASSWORD.Visible = False
Me.CMDOK.Visible = False
Else
Me.TXTADMINPASSWORD.Value = ""
Me.TXTADMINPASSWORD.Visible = False
Me.CMDOK.Visible = False
Me.MultiPage1.Value = 1
End If

End Sub

Private Sub CMDSIGNUP_Click()
Me.TXTADMINPASSWORD.Visible = True
Me.CMDOK.Visible = True
End Sub

Private Sub CommandButton1_Click()

Application.DisplayAlerts = False
ActiveWorkbook.Close SaveChanges:=False
Application.Quit
Application.DisplayAlerts = True


End Sub

Private Sub CommandButton5_Click()
Userform1.MultiPage1.Value = 0
End Sub

Private Sub Label36_Click()
Label36.Visible = False
Me.TXTUSER.SetFocus

End Sub

Private Sub Label37_Click()
Label37.Visible = False
Me.TXTPASSWORD.SetFocus
End Sub



Private Sub ONOFF1_Click()
If Me.ONOFF1.Caption = "OFF" Then
Call ButtonOn1
Else
Call ButtonOff1
End If

End Sub

Private Sub ONOFF2_Click()
If Me.ONOFF2.Caption = "OFF" Then
Call ButtonOn2
Else
Call ButtonOff2
End If

End Sub

Private Sub ONOFF3_Click()
If Me.ONOFF3.Caption = "OFF" Then
Call ButtonOn3
Else
Call ButtonOff3
End If

End Sub
Private Sub TXTPASSWORD_Change()
Label37.Visible = False
Me.Notifywrongpassword.Visible = False
End Sub

Private Sub TXTPASSWORD_Enter()
Label37.Visible = False
Me.Notifywrongpassword.Visible = False
End Sub

Private Sub TXTPASSWORD_Exit(ByVal Cancel As MSForms.ReturnBoolean)
If Me.TXTPASSWORD.Value = "" Then
Label37.Visible = True
End If
End Sub

Private Sub TXTUSER_Change()
Label36.Visible = False
Me.Notifywrongpassword.Visible = False
End Sub

Private Sub TXTUSER_Enter()
Label36.Visible = False
Me.Notifywrongpassword.Visible = False
End Sub

Private Sub TXTUSER_Exit(ByVal Cancel As MSForms.ReturnBoolean)
If Me.TXTUSER.Value = "" Then
Label36.Visible = True
End If
End Sub

Private Sub UserForm_Initialize()

''Hide excel and login userform UI/UX

'Application.Visible = False
Me.MultiPage1.Value = 0
Me.Frame6.Height = Me.MultiPage1.Height
Me.Frame7.Height = Me.MultiPage1.Height
Me.ONOFF1.BackColor = RGB(219, 42, 89)
Me.ONOFF1.ForeColor = RGB(219, 42, 89)

Me.ONOFF2.BackColor = RGB(219, 42, 89)
Me.ONOFF2.ForeColor = RGB(219, 42, 89)

Me.ONOFF3.BackColor = RGB(219, 42, 89)
Me.ONOFF3.ForeColor = RGB(219, 42, 89)

Me.Line1.BackColor = RGB(56, 66, 66)
Me.Line2.BackColor = RGB(56, 66, 66)
Me.Line3.BackColor = RGB(56, 66, 66)

Me.TXTADMINPASSWORD.Visible = False
Me.CMDOK.Visible = False
Me.TXTCHECKPASSWORD.Visible = False
Me.TXTCHECKUSERNAME.Visible = False
Me.Notifywrongpassword.Visible = False

End Sub
Private Sub ButtonOn1()
Do While ONOFF1.Left < Me.Line1.Width - Me.ONOFF1.Width
ONOFF1.Left = ONOFF1.Left + 0.25
DoEvents
Me.ONOFF1.Caption = "ON"
Me.ONOFF1.BackColor = RGB(62, 89, 222)
Me.ONOFF1.ForeColor = RGB(62, 89, 222)
Loop
End Sub

Private Sub ButtonOff1()
Do While ONOFF1.Left > 0
ONOFF1.Left = ONOFF1.Left - 0.25
DoEvents
Me.ONOFF1.Caption = "OFF"
Me.ONOFF1.BackColor = RGB(219, 42, 89)
Me.ONOFF1.ForeColor = RGB(219, 42, 89)
Me.Line1.BackColor = RGB(56, 66, 66)
Loop
End Sub
Private Sub ButtonOn2()
Do While ONOFF2.Left < Me.Line2.Width - Me.ONOFF2.Width
ONOFF2.Left = ONOFF2.Left + 0.25
DoEvents
Me.ONOFF2.Caption = "ON"
Me.ONOFF2.BackColor = RGB(62, 89, 222)
Me.ONOFF2.ForeColor = RGB(62, 89, 222)
Loop
End Sub

Private Sub ButtonOff2()
Do While ONOFF2.Left > 0
ONOFF2.Left = ONOFF2.Left - 0.25
DoEvents
Me.ONOFF2.Caption = "OFF"
Me.ONOFF2.BackColor = RGB(219, 42, 89)
Me.ONOFF2.ForeColor = RGB(219, 42, 89)
Me.Line2.BackColor = RGB(56, 66, 66)
Loop
End Sub

Private Sub ButtonOn3()
Do While ONOFF3.Left < Me.Line3.Width - Me.ONOFF3.Width
ONOFF3.Left = ONOFF3.Left + 0.25
DoEvents
Me.ONOFF3.Caption = "ON"
Me.ONOFF3.BackColor = RGB(62, 89, 222)
Me.ONOFF3.ForeColor = RGB(62, 89, 222)
Loop
End Sub

Private Sub ButtonOff3()
Do While ONOFF3.Left > 0
ONOFF3.Left = ONOFF3.Left - 0.25
DoEvents
Me.ONOFF3.Caption = "OFF"
Me.ONOFF3.BackColor = RGB(219, 42, 89)
Me.ONOFF3.ForeColor = RGB(219, 42, 89)
Me.Line3.BackColor = RGB(56, 66, 66)
Loop
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
If CloseMode = 0 Then
    Cancel = 1

End If
End Sub
