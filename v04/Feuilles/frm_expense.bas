
Option Explicit
Dim i As Integer


Private Sub btn_calendar_Click()
Call Date_Calendar.SelectedDate(Me.TextBox1)
'Me.TextBox1.Value = Date_Calendar.SelectedDate

End Sub

Private Sub btn_Closefrm_Click()

'Call ShowUI
Application.Visible = True
Unload Me
ActiveWorkbook.Sheets("Database").Select
ActiveSheet.Range("$A$1:$I$167").AutoFilter Field:=3
ActiveSheet.Range("$A$1:$I$167").AutoFilter Field:=4
 'ActiveSheet.ShowAllData

ActiveWorkbook.Sheets("Interface").Select


End Sub

Private Sub btn_delete_record_Click()
    Application.ScreenUpdating = False
    If Selected_List = 0 Then

    MsgBox "No row is selected.", vbOKOnly + vbInformation, "Delete"
    Exit Sub

    End If

    Dim i As VbMsgBoxResult

    i = MsgBox("Do you want to delete the selected record?", vbYesNo + vbQuestion, "Confirmation")

    If i = vbNo Then Exit Sub


    ThisWorkbook.Sheets("database").Rows(Selected_List + 1).Delete

    Call Reset

    MsgBox "Selected record has been deleted.", vbOKOnly + vbInformation, "Deleted"

    ' Autofil Serial Number when record is deleted
'

   'To filldown serialNumbers in database
   Sheets("Database").Select
Range("A2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    Range("A2").Select
    ActiveCell.FormulaR1C1 = "1"
    Range("A3").Select
    ActiveCell.FormulaR1C1 = "2"
    Range("A2:A3").Select
    Selection.AutoFill Destination:=Range("fTransaction[S/N]")
    'To Clear extras
    Range("fTransaction[[#Headers],[Date]]").Select
    Selection.End(xlDown).Offset(1, 0).Select
    Selection.End(xlToLeft).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    Selection.End(xlUp).Select

    Application.ScreenUpdating = True
End Sub

Private Sub btn_edit_data_Click()
If Selected_List = 0 Then

    MsgBox "No row is selected.", vbOKOnly + vbInformation, "Edit"
    Exit Sub

    End If

    'Code to update the value to respective controls


    Me.txt_rownumber.Value = Selected_List + 1

    Me.TextBox1.Value = Format(Me.lst_database.List(Me.lst_database.ListIndex, 1), "DD-MMM-YYYY")

    Me.cbo_Category.Value = Me.lst_database.List(Me.lst_database.ListIndex, 2)

    Me.cbo_expensename = Me.lst_database.List(Me.lst_database.ListIndex, 3)

    Me.txt_amount.Value = Me.lst_database.List(Me.lst_database.ListIndex, 4)

    Me.cbo_location = Me.lst_database.List(Me.lst_database.ListIndex, 5)

    Me.txt_Comment.Value = Me.lst_database.List(Me.lst_database.ListIndex, 6)

    MsgBox "Please make the required changes and click on 'Add Entry' button to update."



End Sub

Private Sub btn_export_Click()

' Sheets("Database").Select
'    Range("A1").Select
'    Range(Selection, Selection.End(xlToRight)).Select
'    Range(Selection, Selection.End(xlDown)).Select
'    Selection.Copy
'    Workbooks.Add
'    ActiveSheet.Paste
'    Cells.Select
'    Cells.EntireColumn.AutoFit
' Unload Me

 Dim msgValue As VbMsgBoxResult

msgValue = MsgBox("Do you want to Export to NewWorkBook?", vbYesNo + vbInformation, "Confirmation")

If msgValue = vbNo Then Exit Sub

Call ExportData
'Application.Visible = True
Unload Me

End Sub

Private Sub btn_full_restore_Click()
Call Maximize_Restore
End Sub

Private Sub btn_full_restore_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Unload Me
End Sub

Private Sub btn_Import_Click()
Unload Me
Application.Visible = True
Call ShowUI
ActiveWorkbook.Sheets("Database").Select
Call Getfile

End Sub

Private Sub btn_reset_Click()
Dim msgValue As VbMsgBoxResult

msgValue = MsgBox("Do you want to refresh to form?", vbYesNo + vbInformation, "Confirmation")

If msgValue = vbNo Then Exit Sub

Call Reset


End Sub

Private Sub btn_update_to_dbs_Click()

'Validating entry for emp. name

    Dim Wk As String
    Dim BM As String
    Dim Ab As String

    Wk = frm_expense.cbo_Category.Value
    BM = frm_expense.cbo_expensename.Value
    Ab = frm_expense.txt_amount.Value

    If Wk = "" Then
    MsgBox "Please enter category of expense !", vbOKOnly + vbCritical, "Error"
    Exit Sub
    ElseIf BM = "" Then
    MsgBox "Please enter expense name!", vbOKOnly + vbCritical, "Error"
    Exit Sub
    ElseIf Ab = "" Then
    MsgBox "Please enter expense amount!", vbOKOnly + vbCritical, "Error"
    Exit Sub
    Else
    Call Add_entry

    End If

ActiveWorkbook.Save

MsgBox "Saved Successully!"

Call Reset

End Sub

Private Sub cbo_Category_Click()

Dim sn As Worksheet
Set sn = ThisWorkbook.Sheets("combobox")

Dim i As Integer
Dim n As Integer

n = Application.WorksheetFunction.Match(Me.cbo_Category.Value, sn.Range("1:1"), 0)

Me.cbo_expensename.Clear

For i = 2 To Application.WorksheetFunction.CountA(sn.Cells(1, n).EntireColumn)

Me.cbo_expensename.AddItem sn.Cells(i, n).Value

Next i

End Sub

Private Sub Category_Change()

Application.ScreenUpdating = False
'FIRSTCODE_TRIAL
'If Me.Category.Value = Me.Category.Value Then
'Sheets("Database").Select
'    Application.Goto Reference:="Recordset"
'    ActiveSheet.Range("$A$1:$I$165").AutoFilter Field:=3, Criteria1:=Me.Category.Value
'
'        Else
'
'        MsgBox "No data found"
'        Exit Sub
'
'        End If

      'SECONDCODE NOHEADERS


'Dim Database(1 To 1000000, 1 To 9)
'Dim My_range As Integer
'Dim Colum As Byte
'On Error Resume Next
'
'Sheet3.Range("C1").AutoFilter Field:=3, Criteria1:=Me.Category.Value
'
'For i = 2 To Sheet3.Range("A1000000").End(xlUp).Row
'If Sheet3.Cells(i, 3) = Me.Category Then
'
'My_range = My_range + 1
'For Colum = 1 To 9
'Database(My_range, Colum) = Sheet3.Cells(i, Colum)
'Next Colum
'End If
'Next i
'
'Me.lst_database.List = Database




Dim sn As Worksheet
Set sn = ThisWorkbook.Sheets("combobox")

Dim i As Integer
Dim n As Integer



n = Application.WorksheetFunction.Match(Me.Category.Value, sn.Range("1:1"), 0)


Me.ExpenseName.Clear

For i = 2 To Application.WorksheetFunction.CountA(sn.Cells(1, n).EntireColumn)

Me.ExpenseName.AddItem sn.Cells(i, n).Value

Next i



If Me.Category.Value = "All" Then On Error Resume Next
Me.lst_database.RowSource = "Database"
Me.lst_database.ColumnHeads = False

'Call FilterData
Application.ScreenUpdating = True

End Sub



Private Sub ClearFilter_Click()

Application.ScreenUpdating = False

If Me.Category.Value = LCase(Me.Category.Value) & "*" Then

ActiveWorkbook.Sheets("Database").Select
 'On Error Resume Next
 ActiveSheet.ShowAllData

 Else

 Me.Category.Value = "All"
 'Me.lst_database.RowSource = "Database"
 ActiveWorkbook.Sheets("Database").Select
ActiveSheet.Range("$A$1:$I$167").AutoFilter Field:=3
ActiveSheet.Range("$A$1:$I$167").AutoFilter Field:=4
 'ActiveSheet.ShowAllData

ActiveWorkbook.Sheets("Interface").Select

 Exit Sub


    End If
     Application.ScreenUpdating = True
End Sub



Private Sub ClearFilter2_Click()

Application.ScreenUpdating = False

If Me.Category.Value = LCase(Me.Category.Value) & "*" Then

ActiveWorkbook.Sheets("Database").Select
 'On Error Resume Next
 ActiveSheet.ShowAllData

 Else

 Me.Category.Value = "All"
 'Me.lst_database.RowSource = "Database"
 ActiveWorkbook.Sheets("Database").Select
ActiveSheet.Range("$A$1:$I$167").AutoFilter Field:=3
ActiveSheet.Range("$A$1:$I$167").AutoFilter Field:=4
 'ActiveSheet.ShowAllData

ActiveWorkbook.Sheets("Interface").Select

 Exit Sub


    End If
     Application.ScreenUpdating = True
End Sub

Private Sub ExpenseName_Change()
Call FilterData
End Sub

Private Sub Label17_Click()
Label17.Visible = False
Me.txt_searchrecord.SetFocus
End Sub

Private Sub lst_database_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
If Selected_List = 0 Then

    MsgBox "No row is selected.", vbOKOnly + vbInformation, "Edit"
    Exit Sub

    End If

    'Code to update the value to respective controls

    Me.txt_rownumber.Value = Selected_List + 1

    Me.TextBox1.Value = Format(Me.lst_database.List(Me.lst_database.ListIndex, 1), "DD-MMM-YYYY")

    Me.cbo_Category.Value = Me.lst_database.List(Me.lst_database.ListIndex, 2)

    Me.cbo_expensename = Me.lst_database.List(Me.lst_database.ListIndex, 3)

    Me.txt_amount.Value = Me.lst_database.List(Me.lst_database.ListIndex, 4)

    Me.cbo_location = Me.lst_database.List(Me.lst_database.ListIndex, 5)

    Me.txt_Comment.Value = Me.lst_database.List(Me.lst_database.ListIndex, 6)


    MsgBox "Please make the required changes and click on 'Add Entry' button to update."

    Application.ScreenUpdating = True

End Sub
Private Sub FilterData()
Application.ScreenUpdating = True
Dim Region As String
Dim Item_Type As String
Dim myDB As Range


With Me
'If .Category.ListIndex < 0 Or .ExpenseName.ListIndex < 0 Then Exit Sub
Region = .Category.Value
Item_Type = .ExpenseName.Value
End With
With ActiveWorkbook.Sheets("Database")
Set myDB = .Range("A1:i1").Resize(.Cells(.Rows.Count, 1).End(xlUp).Row)
End With
With myDB
'.AutoFilter 'remove filters
.AutoFilter Field:=3, Criteria1:=Region ' filter data
.SpecialCells(xlCellTypeVisible).AutoFilter Field:=4, Criteria1:=Item_Type ' filter data again
Call UpdateListBox(Me.lst_database, myDB, 1)
'.AutoFilter
End With

Application.ScreenUpdating = False

End Sub
Sub UpdateListBox(lst_database As MSForms.ListBox, myDB As Range, columnToList As Long)

Application.ScreenUpdating = False

Dim Cell As Range, dataValues As Range

 If Category.Value = "All" Then

lst_database.RowSource = "Database"
ElseIf myDB.SpecialCells(xlCellTypeVisible).Count > myDB.Columns.Count Then
    Set dataValues = myDB.Resize(myDB.Rows.Count + 1)

    lst_database.RowSource = ""
   'lst_database.Clear ' we clear the listbox before adding new elements
    For Each Cell In dataValues.Columns(columnToList).SpecialCells(xlCellTypeVisible)
        With Me.lst_database

        If Me.lst_database.RowSource = "Database" Then lst_database.RowSource = ""
        On Error Resume Next
        .AddItem Cell.Value
        'On Error Resume Next
        .List(.ListCount - 1, 1) = Cell.Offset(0, 1).Value
        .List(.ListCount - 1, 2) = Cell.Offset(0, 2).Value
        .List(.ListCount - 1, 3) = Cell.Offset(0, 3).Value
        .List(.ListCount - 1, 4) = Cell.Offset(0, 4).Value
        .List(.ListCount - 1, 5) = Cell.Offset(0, 5).Value
        .List(.ListCount - 1, 6) = Cell.Offset(0, 6).Value
        .List(.ListCount - 1, 7) = Cell.Offset(0, 7).Value
        .List(.ListCount - 1, 8) = Cell.Offset(0, 8).Value
        .List(.ListCount - 1, 9) = Cell.Offset(0, 9).Value
        End With


    Next Cell

Else:

ActiveWorkbook.Sheets("Database").Select

ActiveSheet.ShowAllData
ActiveWorkbook.Sheets("Interface").Select
lst_database.RowSource = ""



End If

lst_database.SetFocus

Application.ScreenUpdating = True

End Sub

Private Sub txt_searchrecord_Change()
Me.Label17.Visible = False

Application.ScreenUpdating = False
Dim Database(1 To 1000000, 1 To 9)
Dim My_range As Integer
Dim Colum As Byte
On Error Resume Next

Me.lst_database.RowSource = ""


For i = 1 To Sheet3.Range("A1000000").End(xlUp).Row

Sheet3.Range("C1").AutoFilter Field:=3, Criteria1:=LCase(Me.txt_searchrecord.Value) & "*"

If LCase(Sheet3.Cells(i, 3)) Like LCase(Me.txt_searchrecord.Value) & "*" = True Then
My_range = My_range + 1

For Colum = 1 To 9
Database(My_range, Colum) = Sheet3.Cells(i, Colum)

Next Colum
End If
Next i

Me.lst_database.List = Database
Me.lst_database.ColumnHeads = False
If Me.txt_searchrecord.Value = "" Then
Me.lst_database.RowSource = "Database"
lst_database.SetFocus
Me.lst_database.ColumnHeads = False
End If


'With lst_database
'
'Me.lst_database.RowSource = ""
'Me.lst_database.AddItem
'        Me.lst_database.List(i - 1, 0) = Cells(i, 1).Value
'        Me.lst_database.List(i - 1, 1) = Cells(i, 2).Value
'        Me.lst_database.List(i - 1, 2) = Cells(i, 3).Value
'        Me.lst_database.List(i - 1, 3) = Cells(i, 4).Value
'        Me.lst_database.List(i - 1, 4) = Cells(i, 5).Value
'        Me.lst_database.List(i - 1, 5) = Cells(i, 6).Value
'        Me.lst_database.List(i - 1, 6) = Cells(i, 7).Value
'        Me.lst_database.List(i - 1, 7) = Cells(i, 8).Value
'        Me.lst_database.List(i - 1, 8) = Cells(i, 9).Value
'
'End With
'
'
'
'
'Else
''Me.lst_database.RowSource = ""
'
'End If
'Next i

Application.ScreenUpdating = True

End Sub

Private Sub txt_searchrecord_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Me.Label17.Visible = False
End Sub

Private Sub txt_searchrecord_DropButtonClick()
Me.Label17.Visible = False
End Sub

Private Sub txt_searchrecord_Enter()
If Me.txt_searchrecord.Value = "" Then
Me.Label17.Visible = True
End If
End Sub
Private Sub txt_searchrecord_Exit(ByVal Cancel As MSForms.ReturnBoolean)
If Me.txt_searchrecord.Value = "" Then
Label17.Visible = True
End If
End Sub

Private Sub UserForm_Activate()
''Reset controls and insert combobox values

Call Reset

Dim sn As Worksheet
Set sn = ThisWorkbook.Sheets("combobox")

Dim p As Integer

Me.cbo_Category.Clear
For p = 1 To Application.WorksheetFunction.CountA(sn.Range("1:1"))
Me.cbo_Category.AddItem sn.Cells(1, p).Value

Next p


End Sub


Private Sub UserForm_Initialize()



'hide userform titleBar
HideBar Me

Application.Visible = False
'ToHide excel worksheet Menus
With Application
    .ScreenUpdating = False
    .Calculation = xlCalculationManual
    .ExecuteExcel4Macro "Show.Toolbar(""Ribbon"",false)"
    .WindowState = xlMaximized
    .CommandBars("Full Screen").Visible = False
    .CommandBars("Worksheet Menu Bar").Enabled = False
    .DisplayStatusBar = False
    .DisplayFormulaBar = False
    .DisplayScrollBars = False
    '.Width = 900
    '.Height = 500

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

frm_expense.TextBox1.Value = Format([Now()], "DD-MMM-YYYY")

Call Reset

Me.LBUSER.Caption = Sheet10.Range("A12").Value

Application.WindowState = xlMaximized

    Me.Top = 0
    Me.Left = 0
    Me.Height = Application.Height - 20
    Me.Width = Application.Width - 5

    With Me.cbo_location

    .AddItem "Clinic Sales"
    .AddItem "Diagnostics Sales (Lab)"
    .AddItem "Inpatient Sales"
    .AddItem "General Sales"
    .AddItem "Store Consumable Sales"

    End With


'Add all Category of expense into filter
Dim Objcell As Range

If Me.Category.Value = "All" Then
Sheets("Sheet3").Select
 ActiveSheet.ShowAllData

 Else

 Me.Category.Value = "All"


End If

For Each Objcell In Range("ExpenseCategory")
Me.Category.AddItem Objcell.Value
Next Objcell


End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
If CloseMode = 0 Then
Cancel = 1

End If

End Sub
