Attribute VB_Name = "create_order_template"
Sub create_order_template()
    Dim order_number As String
    Dim order_date As String
    Dim readiness As String
    
    order_number = InputBox("¬ведите номер заказа")
    order_date = InputBox("¬ведите дату заказа")
    readiness_date = InputBox("¬ведите желаемую дату готовности")
    
    With ThisWorkbook.Worksheets(1)
        Cells(1, 1).Value = "Order є"
        Cells(2, 1).Value = "Order date"
        Cells(3, 1).Value = "Readiness date"
        Cells(1, 2).Value = order_number
        Cells(2, 2).Value = order_date
        Cells(3, 2).Value = readiness_date
        Cells(4, 1).Value = "Confirmed readiness date by supplier"
        
        Cells(5, 1).Value = "Article"
        Cells(5, 2).Value = "Photo"
        Cells(5, 3).Value = "Gender"
        Cells(5, 4).Value = "Color"
        Cells(5, 30).Value = "EXW"
        Cells(5, 31).Value = "Order"
        .Name = order_number
    End With
        
    Dim i As Long
    Dim cellcheck As Range
    
    i = 24
    
    For Each cellcheck In Range("E5:AC5")
        cellcheck.Value = i
        i = i + 1
    Next cellcheck
         
    With ThisWorkbook.Worksheets(1).Range("A:AE")
        .Font.Bold = True
        .Font.Color = 0
        .Font.Size = 16
        .Font.Name = "Calibri"
        .EntireColumn.AutoFit
        .HorizontalAlignment = xlCenter
    End With
        
    Range("A5:AE5").Interior.Color = 15917529
    Range("A1:A4").Font.Color = 1137094
    
    With ThisWorkbook.Worksheets(1).Range("E4:AC4")
        .MergeCells = True
        .Value = "Sizes"
        .Interior.Color = 15917529
    End With
        
    Range("B6:B1048576").ColumnWidth = 25
    Range("B6:B1048576").RowHeight = 85
        
End Sub
