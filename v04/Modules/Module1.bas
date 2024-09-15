Attribute VB_Name = "Module1"
Option Explicit
Public iWidth As Integer
Public iHeight As Integer
Public iLeft As Integer
Public iTop As Integer
Public bState As Boolean

Sub Openme()

frm_expense.Show

End Sub
Sub ShowUI()

'' Showing excel ribbon n menu if a helper cell value is 0 or 1

Application.Calculation = xlCalculationManual
Application.ScreenUpdating = False

Dim Cell As Range

Set Cell = Sheet5.Range("U1")

If Cell.Value = 0 Then

With Application
    .ExecuteExcel4Macro "Show.Toolbar(""Ribbon"", True)"
    .DisplayFormulaBar = True
    .DisplayStatusBar = True
End With

With ActiveWindow
    .DisplayWorkbookTabs = True
    .DisplayHorizontalScrollBar = True
    .DisplayVerticalScrollBar = True
    .DisplayRuler = True
    .DisplayHeadings = True
End With

Sheet5.Shapes("Picture 47").Visible = msoFalse
Sheet5.Shapes("Picture 46").Visible = msoCTrue
Sheet7.Shapes("Picture 17").Visible = msoFalse
Sheet7.Shapes("Picture 18").Visible = msoCTrue

Cell.Value = 1

End If
    
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic

End Sub
Sub HideExcelMenus()

'' hiding excel ribbon n menu if a helper cell value is 0 or 1

Dim Cell As Range

Set Cell = Sheet5.Range("u1")

If Cell.Value = 1 Then

 
With Application
    .ScreenUpdating = False
    .Calculation = xlCalculationManual
    .ExecuteExcel4Macro "Show.Toolbar(""Ribbon"",false)"
    .WindowState = xlMaximized
    .CommandBars("Full Screen").Visible = False
    .CommandBars("Worksheet Menu Bar").Enabled = False
    .DisplayStatusBar = False
    .DisplayFormulaBar = False
    '.DisplayScrollBars = False
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
    


Sheet5.Shapes("Picture 46").Visible = msoFalse
Sheet5.Shapes("Picture 47").Visible = msoCTrue
Sheet7.Shapes("Picture 18").Visible = msoFalse
Sheet7.Shapes("Picture 17").Visible = msoCTrue

Cell.Value = 0
      
End If
End Sub

Sub ShowSYSTEMADMIN()

Systemadmin.Show

End Sub



' To Reset userform controls and connect the listbox to database sheet
'


Sub Reset()
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    
    With frm_expense
    
    '.TextBox1.Value = ""
    .txt_amount.Value = ""

    .txt_Comment.Value = ""
    

    .cbo_Category.Value = ""
    .cbo_expensename.Value = ""
    .cbo_location.Value = ""
    

    .txt_searchrecord.Value = ""
        
    .txt_rownumber.Value = ""
    
    End With
    
    ThisWorkbook.Sheets("Database").Range("E:E").NumberFormat = "[$-ha-Latn-NG] #,##0"
    
    frm_expense.TextBox1.Value = Format([Now()], "DD-MMM-YYYY")
     
    

    Dim irow As Long
    irow = [counta(Database!A:A)] + 1 'identifying the last row
    
    With frm_expense
    
    .lst_database.ColumnCount = 9
    .lst_database.ColumnHeads = True
    .lst_database.ColumnWidths = "25,70,85,85,50,95,105,85,85,"

    End With
    

    If irow > 1 Then
        frm_expense.lst_database.RowSource = "Database!A2:I" & irow
    Else

    frm_expense.lst_database.RowSource = "Database!A2:I2"

    End If

   
 
    
    'Application.ThisWorkbook.Sheets("interface").Cells(7, 2).BackColor = vbWhite
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

End Sub


'' To add value on controls to the database sheet and format fields

Sub Add_entry()

Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
    
    Dim Sh As Worksheet
    Dim irow As Long
    
    Set Sh = ThisWorkbook.Sheets("Database")
    
    If frm_expense.txt_rownumber.Value = "" Then
    
    irow = [counta(Database!A:A)] + 1
    
    Else
    
        irow = frm_expense.txt_rownumber.Value
        
        End If
        
    
    
    With Sh
    
    .Cells(irow, 1) = irow - 1
    
    .Cells(irow, 2) = frm_expense.TextBox1.Value
    
    .Cells(irow, 3) = frm_expense.cbo_Category.Value
    
    .Cells(irow, 4) = frm_expense.cbo_expensename.Value
    
    .Cells(irow, 5) = frm_expense.txt_amount.Value
    
    .Cells(irow, 6) = frm_expense.cbo_location.Value
    
    .Cells(irow, 7) = frm_expense.txt_Comment
    
    .Cells(irow, 8) = Sheet10.Range("A15").Value 'USERid on Entry
    
    .Cells(irow, 9) = [text(now(),"DD-MM-YY HH:MM:SS Am/Pm")]
    
    'Formating the columns
    
    Sh.Range("B:B").NumberFormat = "dd-mm-yyyy"
    Sh.Range("A:A").NumberFormat = "0"
    Sh.Range("E:E").NumberFormat = "0.00"
    'sh.Range("E:E").NumberFormat = "0.00"

    End With
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

End Sub


Function Selected_List() As Long

'helperfunction for editing data visualized in listbox via ID

Dim i As Long
Selected_List = 0

For i = 0 To frm_expense.lst_database.ListCount - 1

    If frm_expense.lst_database.Selected(i) = True Then

        Selected_List = i + 1
        Exit For

        End If


Next i
End Function

Sub Maximize_Restore()

''To maximize the userform

Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual

Dim iWidth As Integer
Dim iHeight As Integer
Dim iLeft As Integer
Dim iTop As Integer
Dim bState As Boolean

    If Not bState = True Then
    
        iWidth = frm_expense.Width
        iHeight = frm_expense.Height
        iTop = frm_expense.Top
        iLeft = frm_expense.Left
        
        'Code for full screen
        
        With Application
            
            .WindowState = xlMaximized
            frm_expense.Zoom = Int(.Width / frm_expense.Width * 100)
            
           frm_expense.StartUpPosition = 0
            frm_expense.Left = .Left
            frm_expense.Top = .Top
            frm_expense.Width = .Width
            frm_expense.Height = .Height
            
        End With
        
        frm_expense.btn_full_restore.Caption = "Close"
        bState = True
    
    Else
    
        With Application

            .WindowState = xlMaximized
            frm_expense.Zoom = 100
            frm_expense.StartUpPosition = 0
            frm_expense.Left = iLeft
            frm_expense.Width = iWidth
            frm_expense.Height = iHeight
            frm_expense.Top = iTop

        End With
        
        frm_expense.btn_full_restore.Caption = "Full Screen"
        Unload frm_expense
        
        bState = False
        
        
        
    End If


Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic

End Sub


   



Sub ExportData()

''Exporting the database sheet

Sheets("Database").Select
    Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Workbooks.Add
    ActiveSheet.Paste
    Cells.Select
    Cells.EntireColumn.AutoFit
 
 
 
 
 
 
 
End Sub

 Sub Getfile()

''requiring a file path inother to import the file to the database sheet

Dim FileSelect As Variant
Dim wb As Workbook
Dim i As Integer

On Error GoTo errHandler:


Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual

FileSelect = Application.GetOpenFilename(filefilter:="Excel Files, *.xl*", MultiSelect:=False)

If FileSelect = False Then
    
    MsgBox "Select the file name", vbCritical
    ActiveWorkbook.Sheets("Interface").Select
    Call HideExcelMenus
    Exit Sub
End If

Sheet3.Range("J21").Value = FileSelect
Sheet3.Range("J22").Value = "Sheet 1"
Sheet3.Range("J23").Value = "Database"

'Sheets(Sheet3).Shapes.filepath.TextFrame.Characters.Text = "Wait"
Set wb = Workbooks.Open(FileSelect)
Sheet3.Range("N4:N100").ClearContents
'For i = 1 To Sheets.Count
'    Sheet3.Range ("N") & 1 + 3 = Sheets(i).Name
'
'Next i

wb.Close False
Application.ScreenUpdating = True
Application.Calculation = xlCalculationManual

Exit Sub
'
errHandler:
    MsgBox "An Error has Occurred " & vbCrLf & "The error number is: " _
    & Err.Number & vbCrLf & Err.Description & vbCrLf & _
    "Please Notify the administrator"

End Sub
Public Sub GetRange()

Dim FileSelect As Variant
Dim wb As Workbook
Dim Addme As Range, _
    CopyData As Range, _
    Bk As Range, _
    Sh As Range, _
    Tb As Range, _
    C As Range
On Error GoTo errHandler:

Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual

Sheet3.Select
    If Range("A2").Value = "" Then
    ActiveSheet.ListObjects("fTransaction").Resize Range("$A$1:$I$2")
    Else
    End If

For Each C In Sheet3.Range("J21")
    If C.Value = "" Then
    MsgBox "You have to provide an excel file location in other to import", vbCritical
    Exit Sub
    End If
    Next C

    
    Set Bk = Sheet3.Range("J21")
'File path of book to import From
    Set Sh = Sheet3.Range("J22")
'Sheet to import
    Set Tb = Sheet3.Range("J23")
'sheet in this workbook to send it to set the destination
    Set Addme = Worksheets(Tb.Value).Range("A" & Rows.Count).End(xlUp)
'Open the workbook
    Set wb = Workbooks.Open(Bk)
'Set the copy range
    Range("A2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
'Copy and Paste the data
    Selection.Copy
    Addme.PasteSpecial xlPasteValues
'Clear the clipboard
    Application.CutCopyMode = False
'Close the Workbook
    wb.Close False
'To filldown serialNumbers in database
    Sheet3.Select
    Range("J21").Clear
 
 'To filldown serialNumbers in database
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
    
'to format date
    Range("fTransaction[Date]").Select
    Selection.NumberFormat = "m/d/yyyy"
    
'return to the interface sheet and hideMenus
    Sheet5.Select
    Call HideExcelMenus
    frm_expense.Show
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
    Exit Sub

errHandler:
    MsgBox "An Error has Occurred " & vbCrLf & "The error numbe is: " _
    & Err.Number & vbCrLf & Err.Description & vbCrLf & _
    "Please Notify the administrator"

End Sub


Sub GetSheets()

Sheet3.Range("O4:O100").ClearContents
Dim i As Integer
For i = 1 To Sheets.Count
    Sheet3.Range("O" & 1 + 3) = Sheets(i).Name
Next i
End Sub
