Attribute VB_Name = "Module4"
Sub refreshDashboard()
'attached to the dashboard button to refresh the visualizations

Application.ActiveWorkbook.RefreshAll

End Sub

Sub RefreshPivotTableAnalysis()


' refresh the pivotables on sheet initiation
 
    ActiveSheet.PivotTables("PivotTable1").PivotCache.Refresh
    
End Sub


