VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Private Sub Analysis2_Click()
Sheet7.Select
ActiveWindow.DisplayHeadings = False

End Sub

Private Sub Dashboard2_Click()
Sheet9.Select
End Sub

Private Sub Interface2_Click()
Sheet5.Select
ActiveWindow.DisplayHeadings = False

End Sub

Private Sub SysAdmin2_Click()
Call ShowSYSTEMADMIN
End Sub


Private Sub Worksheet_Activate()

'ActiveSheet.Shapes.Range(Array("Rectangle 19", "Rectangle 20", "Group 11")).Select
    'ActiveSheet.Protect DrawingObjects:=True, Contents:=False, Scenarios:= _
        False

       
    
     
End Sub

