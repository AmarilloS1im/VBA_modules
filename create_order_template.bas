Attribute VB_Name = "create_order_template"
Sub create_order_template()
    Dim order_number As String
    Dim order_date As String
    Dim readiness As String
    Dim article As Range
    
    order_number = InputBox("Введите номер заказа")
    order_date = InputBox("Введите дату заказа")
    readiness_date = InputBox("Введите желаемую дату готовности")
    
    With ThisWorkbook.ActiveSheet
        Cells(1, 1).Value = "Order №"
        Cells(2, 1).Value = "Order date"
        Cells(3, 1).Value = "Readiness date"
        Cells(4, 1).Value = "Confirmed readiness date by supplier"
        
        Cells(1, 2).Value = order_number
        Cells(2, 2).Value = order_date
        Cells(3, 2).Value = readiness_date
        Cells(4, 2).Value = "?? ?? ????"
        

        
        Cells(5, 1).Value = "Article"
        Cells(5, 2).Value = "Model"
        Cells(5, 3).Value = "Color"
        Cells(5, 4).Value = "ART № Rehard"
        Cells(5, 5).Value = "Model Rehard"
        Cells(5, 6).Value = "Color Rehard"
        Cells(5, 7).Value = "Gender"
        Cells(5, 8).Value = "Photo"
        Cells(5, 9).Value = "UP material"
        Cells(5, 10).Value = "Lining"
        Cells(5, 11).Value = "Insole"
        Cells(5, 12).Value = "Outsole"
        Cells(5, 13).Value = "MOQ"
        
        Cells(5, 45).Value = "EXW Price"
        Cells(5, 46).Value = "Order"
        .Name = order_number
    End With
        
    Dim i As Long
    Dim cellcheck As Range
    
    i = 19
    
    For Each cellcheck In Range("N5:AR5")
        cellcheck.Value = i
        i = i + 1
    Next cellcheck
         
    With ThisWorkbook.ActiveSheet.Range("A:AT")
        .Font.Bold = True
        .Font.Color = 0
        .Font.Size = 16
        .Font.Name = "Calibri"
        .EntireColumn.AutoFit
        .HorizontalAlignment = xlCenter
    End With
        
    Range("A5:AT5").Interior.Color = 15917529
    Range("A1:A4").Font.Color = 1137094
    Range("A1:A4").Select
    
    Range("B4").Font.Color = 192
    
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    
    
    Range("D5:F5").Font.Color = 192
    
    With ThisWorkbook.ActiveSheet.Range("N4:AR4")
        .MergeCells = True
        .Value = "Sizes"
        .Interior.Color = 15917529
    End With
        
    Range("B6:B1048576").ColumnWidth = 25
    Range("B6:B1048576").RowHeight = 85
    
    Rows("5:5").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With

    
    
    
    
    
    
    'Dim objCloseBook As Workbook
    
    
    'MsgBox "Откойте 'Рабочий файл' в нужной папке" & vbCrLf & "*****************************************"
    'Открываем рабочий файл
    'order_file = Application.GetOpenFilename("Excel files(*.xls*),*.xls*", 1, "Выбрать Excel файлы", , False)
    
    'Отключаем обновление экрана
    'Application.ScreenUpdating = False
    'Set objCloseBook = Workbooks.Open(order_file)
    'If VarType(order_file) = vbBoolean Then
        'Была нажата кнопка отмены-выход из процедуры
        'Exit Sub
    'End If
    'i = 6
    'For Each article In objCloseBook.Worksheets(1).Range("I2:I148")
        'Workbooks("1.xlsx").Worksheets(1).Cells(i, 2) = article.Value
        'i = i + 1
    'Next article
    
    
    'objCloseBook.Close False

    
    
    'Application.ScreenUpdating = True
    
    
        
End Sub
