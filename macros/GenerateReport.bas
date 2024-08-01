Attribute VB_Name = "Module4"
Sub GenerateReport()
    Dim wsInventory As Worksheet
    Dim wsReport As Worksheet
    Dim rngInventory As Range
    Dim lastReportRow As Long
    
    ' Set Worksheets
    Set wsInventory = Sheets("Inventory")
    Set wsReport = Sheets("Report")
    
    ' Clear the old report
    wsReport.Cells.Clear
    
    ' Report headers
    wsReport.Cells(1, 1).Value = "Item Code"
    wsReport.Cells(1, 2).Value = "Item Name"
    wsReport.Cells(1, 3).Value = "Quantity in Stock"
    wsReport.Cells(1, 4).Value = "Minimum Level"
    
    ' Define the data range and the last line of the report
    Set rngInventory = wsInventory.Range("A2:A" & wsInventory.Cells(wsInventory.Rows.Count, 1).End(xlUp).Row)
    lastReportRow = 2
    
    ' Report generation
    For Each cellInventory In rngInventory
        wsReport.Cells(lastReportRow, 1).Value = cellInventory.Value
        wsReport.Cells(lastReportRow, 2).Value = cellInventory.Offset(0, 1).Value
        wsReport.Cells(lastReportRow, 3).Value = cellInventory.Offset(0, 2).Value
        wsReport.Cells(lastReportRow, 4).Value = cellInventory.Offset(0, 3).Value
        lastReportRow = lastReportRow + 1
    Next cellInventory
End Sub

