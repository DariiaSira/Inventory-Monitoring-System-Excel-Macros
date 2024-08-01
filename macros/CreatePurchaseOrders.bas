Attribute VB_Name = "Module3"
Sub CreatePurchaseOrders()
    Dim wsInventory As Worksheet
    Dim wsOrders As Worksheet
    Dim rngInventory As Range
    Dim cellInventory As Range
    Dim lastOrderRow As Long
    Dim requiredQuantity As Double
    Dim orderQuantity As Double
    
    ' Set worksheets
    Set wsInventory = Sheets("Inventory")
    Set wsOrders = Sheets("Purchase Orders")
    
    ' Clear old report
    wsOrders.Cells.Clear
    
    ' Headers of the new report
    wsOrders.Cells(1, 1).Value = "Item Code"
    wsOrders.Cells(1, 2).Value = "Item Name"
    wsOrders.Cells(1, 3).Value = "Required Quantity"
    wsOrders.Cells(1, 4).Value = "Order Quantity"
    
    ' Determine the data range and the last line of the order
    Set rngInventory = wsInventory.Range("A2:A" & wsInventory.Cells(wsInventory.Rows.Count, 1).End(xlUp).Row)
    
    ' Last order row in Purchase Orders
    lastOrderRow = 2
    
    ' Creating purchase orders
    For Each cellInventory In rngInventory
        ' Definition of required and ordered quantities
        requiredQuantity = cellInventory.Offset(0, 3).Value
        orderQuantity = requiredQuantity - cellInventory.Offset(0, 2).Value
        
        ' If replenishment is required
        If cellInventory.Offset(0, 2).Value < cellInventory.Offset(0, 3).Value Then
            wsOrders.Cells(lastOrderRow, 1).Value = cellInventory.Value
            wsOrders.Cells(lastOrderRow, 2).Value = cellInventory.Offset(0, 1).Value
            wsOrders.Cells(lastOrderRow, 3).Value = requiredQuantity
            wsOrders.Cells(lastOrderRow, 4).Value = orderQuantity
            lastOrderRow = lastOrderRow + 1
        End If
    Next cellInventory
End Sub

