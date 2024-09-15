Attribute VB_Name = "Module2"

Option Explicit

'' interface sheet clock (Start)

Sub Start_Clock()


Dim Sh As Worksheet
Set Sh = ActiveSheet

Sh.Range("N1").Value = "Start"

x:
VBA.DoEvents
If Sh.Range("N1").Value = "Stop" Then Exit Sub
Application.Calculate
GoTo x


End Sub

'' interface sheet clock (Stop)

Sub Stop_Clock()

Dim Sh As Worksheet
Set Sh = ActiveSheet

Sh.Range("N1").Value = "Stop"

 

End Sub


