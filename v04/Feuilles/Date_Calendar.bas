


Option Explicit

Private Sub CmbMonth_Change()

If Me.CmbMonth.Value <> "" And Me.CmbYear.Value <> "" Then
    Call Show_Dates
    Me.lblSelectedMonth = Me.CmbMonth & "-" & Me.CmbYear
End If

End Sub

Private Sub CmbYear_Change()
    If Me.CmbMonth.Value <> "" And Me.CmbYear.Value <> "" Then
        Call Show_Dates
        Me.lblSelectedMonth = Me.CmbMonth & "-" & Me.CmbYear
    End If
End Sub

Sub ButtonClick(btn As MSForms.CommandButton)
    With btn
        If .Caption <> "" Then
            Me.TextBox1.Value = .Caption & "-" & Left(Me.CmbMonth.Value, 3) & "-" & Me.CmbYear.Value
            Unload Me
        End If

    End With
End Sub

Private Sub CommandButton1_Click()
    Call ButtonClick(Me.CommandButton1)
End Sub
Private Sub CommandButton2_Click()
    Call ButtonClick(Me.CommandButton2)
End Sub
Private Sub CommandButton3_Click()
    Call ButtonClick(Me.CommandButton3)
End Sub
Private Sub CommandButton4_Click()
    Call ButtonClick(Me.CommandButton4)
End Sub
Private Sub CommandButton5_Click()
    Call ButtonClick(Me.CommandButton5)
End Sub
Private Sub CommandButton6_Click()
    Call ButtonClick(Me.CommandButton6)
End Sub
Private Sub CommandButton7_Click()
    Call ButtonClick(Me.CommandButton7)
End Sub
Private Sub CommandButton8_Click()
    Call ButtonClick(Me.CommandButton8)
End Sub
Private Sub CommandButton9_Click()
    Call ButtonClick(Me.CommandButton9)
End Sub
Private Sub CommandButton10_Click()
    Call ButtonClick(Me.CommandButton10)
End Sub
Private Sub CommandButton11_Click()
    Call ButtonClick(Me.CommandButton11)
End Sub
Private Sub CommandButton12_Click()
    Call ButtonClick(Me.CommandButton12)
End Sub
Private Sub CommandButton13_Click()
    Call ButtonClick(Me.CommandButton13)
End Sub
Private Sub CommandButton14_Click()
    Call ButtonClick(Me.CommandButton14)
End Sub
Private Sub CommandButton15_Click()
    Call ButtonClick(Me.CommandButton15)
End Sub
Private Sub CommandButton16_Click()
    Call ButtonClick(Me.CommandButton16)
End Sub
Private Sub CommandButton17_Click()
    Call ButtonClick(Me.CommandButton17)
End Sub
Private Sub CommandButton18_Click()
    Call ButtonClick(Me.CommandButton18)
End Sub
Private Sub CommandButton19_Click()
    Call ButtonClick(Me.CommandButton19)
End Sub
Private Sub CommandButton20_Click()
    Call ButtonClick(Me.CommandButton20)
End Sub
Private Sub CommandButton21_Click()
    Call ButtonClick(Me.CommandButton21)
End Sub
Private Sub CommandButton22_Click()
    Call ButtonClick(Me.CommandButton22)
End Sub
Private Sub CommandButton23_Click()
    Call ButtonClick(Me.CommandButton23)
End Sub
Private Sub CommandButton24_Click()
    Call ButtonClick(Me.CommandButton24)
End Sub
Private Sub CommandButton25_Click()
    Call ButtonClick(Me.CommandButton25)
End Sub
Private Sub CommandButton26_Click()
    Call ButtonClick(Me.CommandButton26)
End Sub
Private Sub CommandButton27_Click()
    Call ButtonClick(Me.CommandButton27)
End Sub
Private Sub CommandButton28_Click()
    Call ButtonClick(Me.CommandButton28)
End Sub
Private Sub CommandButton29_Click()
    Call ButtonClick(Me.CommandButton29)
End Sub
Private Sub CommandButton30_Click()
    Call ButtonClick(Me.CommandButton30)
End Sub
Private Sub CommandButton31_Click()
    Call ButtonClick(Me.CommandButton31)
End Sub
Private Sub CommandButton32_Click()
    Call ButtonClick(Me.CommandButton32)
End Sub
Private Sub CommandButton33_Click()
    Call ButtonClick(Me.CommandButton33)
End Sub
Private Sub CommandButton34_Click()
    Call ButtonClick(Me.CommandButton34)
End Sub
Private Sub CommandButton35_Click()
    Call ButtonClick(Me.CommandButton35)
End Sub
Private Sub CommandButton36_Click()
    Call ButtonClick(Me.CommandButton36)
End Sub
Private Sub CommandButton37_Click()
    Call ButtonClick(Me.CommandButton37)
End Sub
Private Sub CommandButton38_Click()
    Call ButtonClick(Me.CommandButton38)
End Sub
Private Sub CommandButton39_Click()
    Call ButtonClick(Me.CommandButton39)
End Sub
Private Sub CommandButton40_Click()
    Call ButtonClick(Me.CommandButton40)
End Sub
Private Sub CommandButton41_Click()
    Call ButtonClick(Me.CommandButton41)
End Sub
Private Sub CommandButton42_Click()
    Call ButtonClick(Me.CommandButton42)
End Sub

Private Sub img_Next_Click()
    On Error Resume Next
    If Me.CmbMonth.ListIndex = 11 Then
        Me.CmbMonth.ListIndex = 0
        Me.CmbYear.Value = Me.CmbYear.Value + 1
    Else
        Me.CmbMonth.ListIndex = Me.CmbMonth.ListIndex + 1
    End If
End Sub

Private Sub img_previous_Click()
    On Error Resume Next
    If Me.CmbMonth.ListIndex = 0 Then
        Me.CmbMonth.ListIndex = 11
        Me.CmbYear.Value = Me.CmbYear.Value - 1
    Else
        Me.CmbMonth.ListIndex = Me.CmbMonth.ListIndex - 1
    End If
End Sub



Private Sub UserForm_Activate()

Dim i As Integer
Dim Year_Start, Year_End As Integer

'================= Add Months to List ==============
With Me.CmbMonth
    .Clear
    For i = 1 To 12
        .AddItem VBA.Format(VBA.DateSerial(2018, i, 1), "MMMM")
    Next i

    .Value = VBA.Format(VBA.Date, "MMMM")
End With

'================ Add Years =======================

  Year_Start = VBA.Year(VBA.Date) - 20
  Year_End = VBA.Year(VBA.Date) + 20

With Me.CmbYear
    .Clear
    For i = Year_Start To Year_End
        .AddItem i
    Next i

    .Value = VBA.Format(VBA.Date, "YYYY")

End With

Call Show_Dates

If Me.TextBox1.Value <> "" Then
    Call Show_Selected_Date(CDate(Me.TextBox1.Value))
End If



End Sub

Private Sub Show_Dates()


    Dim first_Date As Date
    first_Date = VBA.DateValue("1-" & Me.CmbMonth.Value & "-" & Me.CmbYear.Value)
    Dim last_day As Integer
    last_day = VBA.Day(VBA.DateSerial(VBA.Year(first_Date), VBA.Month(first_Date) + 1, 1) - 1)


    Dim i As Integer
    Dim btn As CommandButton

    '============ Clear All button
    For i = 1 To 42
        Set btn = Me.Controls("CommandButton" & i)
        btn.Caption = ""
    Next i

    '====================
    For i = 1 To 7   'Set first date of month
        Set btn = Me.Controls("CommandButton" & i)

        If VBA.Weekday(first_Date) = i Then
            btn.Caption = "1"
        Else
            btn.Caption = ""
        End If
    Next i

    Dim btn1 As CommandButton
    Dim btn2 As CommandButton

    For i = 1 To 41
        Set btn1 = Me.Controls("CommandButton" & i)
        Set btn2 = Me.Controls("CommandButton" & i + 1)
        If btn1.Caption <> "" Then
            If last_day > btn1.Caption Then
               btn2.Caption = btn1.Caption + 1
            End If
        End If
    Next i

Call Reset_Colors

End Sub

Private Sub Reset_Colors()

Dim i As Integer
Dim btn As CommandButton
Me.img_Star.Visible = False
For i = 1 To 42
    Set btn = Me.Controls("CommandButton" & i)

    With btn
        .BackColor = VBA.RGB(255, 215, 0)  'set background colors
        .Enabled = True  'Enable All

        If .Caption = "" Then  'Disbale for blanks
            .Enabled = False
            .BackColor = VBA.RGB(200, 200, 200)
        End If

    End With

Next i

End Sub

 Function SelectedDate(Target_Control As Object) As String

    Dim str As String

    If (TypeName(Target_Control)) = "TextBox" Or TypeName(Target_Control) = "Range" Then str = Target_Control.Value
    If (TypeName(Target_Control)) = "CommandButton" Or TypeName(Target_Control) = "Label" Then str = Target_Control.Caption

    If IsDate(str) Then
        Me.TextBox1.Value = VBA.Format(CDate(str), "D-MMM-YYYY")
        Else
        Me.TextBox1.Value = ""
    End If

    Me.Show

    If (TypeName(Target_Control)) = "TextBox" Or TypeName(Target_Control) = "Range" Then
         Target_Control.Value = Me.TextBox1.Value
    ElseIf (TypeName(Target_Control)) = "CommandButton" Or TypeName(Target_Control) = "Label" Then
         Target_Control.Caption = Me.TextBox1.Value
    Else
        SelectedDate = Me.TextBox1.Value
    End If



End Function

Sub Show_Selected_Date(dt As Date)
    Dim i As Integer
    Dim btn As MSForms.CommandButton
    On Error Resume Next
    Me.CmbMonth.Value = VBA.Format(dt, "MMMM")
    Me.CmbYear.Value = VBA.Format(dt, "YYYY")

    For i = 1 To 42
        Set btn = Me.Controls("CommandButton" & i)
        If btn.Caption = CStr(VBA.Day(dt)) Then

            Me.img_Star.Left = btn.Left + 3
            Me.img_Star.Top = btn.Top + 3
            Me.img_Star.Visible = True
            btn.BackColor = vbWhite

        End If
    Next i

End Sub
