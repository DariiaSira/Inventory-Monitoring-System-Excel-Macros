Attribute VB_Name = "Module1"
Sub StockUpdate()
    Dim wsInventory As Worksheet
    Dim wsIncoming As Worksheet
    Dim rngInventory As Range
    Dim rngIncoming As Range
    Dim cellInventory As Range
    Dim cellIncoming As Range
    Dim itemCode As String
    Dim quantity As Double
    Dim foundCell As Range
    
    ' Set the worksheets
    Set wsInventory = Sheets("Inventory")
    Set wsIncoming = Sheets("Incoming Shipments")
    Set wsOutgoing = Sheets("Outgoing Shipments")
    
    ' Define the ranges
    Set rngInventory = wsInventory.Range("A2:A" & wsInventory.Cells(wsInventory.Rows.Count, 1).End(xlUp).Row)
    Set rngIncoming = wsIncoming.Range("A2:A" & wsIncoming.Cells(wsIncoming.Rows.Count, 1).End(xlUp).Row)
    Set rngOutgoing = wsOutgoing.Range("A2:A" & wsOutgoing.Cells(wsOutgoing.Rows.Count, 1).End(xlUp).Row)
    
    ' Process incoming shipments
    For Each cellIncoming In rngIncoming
        itemCode = cellIncoming.Value
        quantity = cellIncoming.Offset(0, 2).Value ' Quantity of the incoming shipment
        
        ' Find the corresponding item in the inventory
        Set foundCell = rngInventory.Find(What:=itemCode, LookIn:=xlValues, LookAt:=xlWhole)
        
        If Not foundCell Is Nothing Then
            ' Update the quantity in inventory
            foundCell.Offset(0, 2).Value = foundCell.Offset(0, 2).Value + quantity
        Else
            ' Item not found in inventory, you might want to handle this case
            MsgBox "Item code " & itemCode & " not found in inventory.", vbExclamation
        End If
    Next cellIncoming
    
    ' Process outgoing shipments
    For Each cellOutgoing In rngOutgoing
        itemCode = cellOutgoing.Value
        quantity = cellOutgoing.Offset(0, 2).Value ' Quantity of the outgoing shipment
        
        ' Find the corresponding item in the inventory
        Set foundCell = rngInventory.Find(What:=itemCode, LookIn:=xlValues, LookAt:=xlWhole)
        
        If Not foundCell Is Nothing Then
            ' Update the quantity in inventory
            foundCell.Offset(0, 2).Value = foundCell.Offset(0, 2).Value - quantity
        Else
            ' Item not found in inventory, you might want to handle this case
            MsgBox "Item code " & itemCode & " not found in inventory.", vbExclamation
        End If
    Next cellOutgoing
    
End Sub


