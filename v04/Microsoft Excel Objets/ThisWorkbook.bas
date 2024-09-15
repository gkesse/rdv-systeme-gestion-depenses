VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Sub Refresh()

Application.ThisWorkbook.EnableConnections
Application.ActiveWorkbook.RefreshAll

Call Start_Clock


End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
On Error GoTo errHandler:
ThisWorkbook.Saved = True
Application.DisplayAlerts = False
Cancel = False
ThisWorkbook.Saved = True
ThisWorkbook.Close
'This will prevent any other BeforeClose code from running, if any.
errHandler:
    MsgBox "User Account " & vbCrLf & "Expense Management System: " _
    & Err.Number & vbCrLf & Err.Description & vbCrLf & _
    "Weldone, your work has been saved!"
    

    

'
End Sub

Private Sub Workbook_Open()

'hide excel and show only userform

'On Error GoTo Incorrect
Userform1.Show

If Userform1.Visible = True Then
    'Application.Visible = False
Else
    'Application.Visible = True

End If


Dim Cell As Range
Set Cell = Sheet5.Range("U1")

With Application
    .ScreenUpdating = False
    .Calculation = xlCalculationManual
    .ExecuteExcel4Macro "Show.Toolbar(""Ribbon"",False)"
    .WindowState = xlMaximized
    .CommandBars("Full Screen").Visible = True
    .CommandBars("Worksheet Menu Bar").Enabled = False
    .DisplayStatusBar = False
    .DisplayFormulaBar = False
    .DisplayScrollBars = True
    '.Width = 1000
    '.Height = 800
    
End With

With ActiveWindow
    .DisplayWorkbookTabs = False
    .DisplayRuler = False
    
    .DisplayHeadings = False
End With

With Application
    .ScreenUpdating = True
    .Calculation = xlCalculationAutomatic
    End With
    
    'ShowExcelorHide, to place 0 in cell
    Cell.Value = 0
    Sheet5.Shapes("Picture 46").Visible = msoFalse
    Sheet5.Shapes("Picture 47").Visible = msoCTrue
    
  

'Incorrect:
'Call MsgBox("Welcome Back", vbInformation, "Admin")

End Sub

