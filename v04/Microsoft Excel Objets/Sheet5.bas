VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Analysis_Click()
Sheet7.Select

ActiveWindow.DisplayHeadings = False

End Sub

Private Sub Dashboard_Click()
Sheet9.Select
End Sub

Private Sub Goto_DBS_Click()
ThisWorkbook.Sheets("Database").Activate
'ActiveWorkbook.Sheets("Database").Select

End Sub

Private Sub Interface_Click()
Sheet5.Select

ActiveWindow.DisplayHeadings = False

End Sub


Private Sub SysAdmin_Click()
Call ShowSYSTEMADMIN
End Sub

Private Sub Worksheet_Activate()
 
 
' ActiveSheet.Shapes.Range(Array("Rectangle 2", "Rectangle 4", "Rectangle 13", "Rectangle 17" _
        , "Rectangle 18", "Group 28", "Rounded Rectangle 15", "Rounded Rectangle 19", _
        "Picture 36", "Picture 35", "Picture 37", "Picture 39", "Picture 40", _
        "Picture 41", "Picture 42", "Picture 43", "Picture 44", "Picture 45", _
        "Picture 45", "Freeform: Shape 1")).Select

    'ActiveSheet.Protect DrawingObjects:=True, Contents:=False, Scenarios:= _
        False
    
    
 
    
End Sub

