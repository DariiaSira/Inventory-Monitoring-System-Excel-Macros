Attribute VB_Name = "Module2"
Sub CheckStockLevels()
    Dim wsInventory As Worksheet
    Dim wsMessages As Worksheet
    Dim rngInventory As Range
    Dim cellInventory As Range
    Dim lastMessageRow As Long
    
    ' Set worksheet
    Set wsInventory = Sheets("Inventory")
    
    ' Create or clear a message worksheet
    On Error Resume Next
    Set wsMessages = Sheets("Stock Alerts")
    On Error GoTo 0
    If wsMessages Is Nothing Then
        Set wsMessages = Sheets.Add(After:=Sheets(Sheets.Count))
        wsMessages.Name = "Stock Alerts"
    Else
        wsMessages.Cells.Clear
    End If
    
    ' Set headers on the message sheet
    wsMessages.Cells(1, 1).Value = "Item Code"
    wsMessages.Cells(1, 2).Value = "Item Name"
    wsMessages.Cells(1, 3).Value = "Current Quantity"
    wsMessages.Cells(1, 4).Value = "Minimum Level"
    wsMessages.Cells(1, 5).Value = "Status"
    
    ' Define the data range
    Set rngInventory = wsInventory.Range("A2:A" & wsInventory.Cells(wsInventory.Rows.Count, 1).End(xlUp).Row)
    lastMessageRow = 2 ' Initial line for messages
    
    ' Check stock levels
    For Each cellInventory In rngInventory
        If cellInventory.Offset(0, 2).Value < cellInventory.Offset(0, 3).Value Then
            ' Write low stock level information to the message sheet
            wsMessages.Cells(lastMessageRow, 1).Value = cellInventory.Value ' Item Code
            wsMessages.Cells(lastMessageRow, 2).Value = cellInventory.Offset(0, 1).Value ' Product Name
            wsMessages.Cells(lastMessageRow, 3).Value = cellInventory.Offset(0, 2).Value ' Current quantity
            wsMessages.Cells(lastMessageRow, 4).Value = cellInventory.Offset(0, 3).Value ' Minimum level
            wsMessages.Cells(lastMessageRow, 5).Value = "Needs Restocking" ' Status
            
            lastMessageRow = lastMessageRow + 1 ' Go to next line
            
        End If
    Next cellInventory
End Sub

