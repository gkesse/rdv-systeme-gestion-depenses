VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Analysis3_Click()
Sheet7.Select
End Sub

Private Sub Dashboard3_Click()
Sheet9.Select
End Sub

Private Sub Interface3_Click()
Sheet5.Select
End Sub

Private Sub SysAdmin3_Click()
Call ShowSYSTEMADMIN
End Sub

Private Sub Worksheet_Activate()

'ActiveSheet.Shapes.Range(Array("Rectangle 83", "Rectangle 84", "Group 80", "Freeform: Shape 2", "Rectangle 10")) _
    .Select
     'ActiveSheet.Protect DrawingObjects:=True, Contents:=False, Scenarios:= _
        False
       
  

End Sub

