Attribute VB_Name = "create_TNVD_template"
'Этот модь создает шаблон файла для согласование ТНВЭД кодов
Sub create_TNVD_template()
    Dim array_file() As String
    Dim i As Integer
    Dim file_name As String
    
    array_file = Split(ThisWorkbook.Name, ".")
    
    For i = 0 To UBound(array_file) - 1
        file_name = file_name + array_file(i) + "."
    Next i
    
    file_name = Left(file_name, (Len(file_name) - 1))
    
    
    With ThisWorkbook.Worksheets(1)
        Cells(1, 1).Value = "АРТИКУЛ КАК У ПРОИЗВОДИТЕЛЯ"
        Cells(1, 2).Value = "КАТЕГОРИЯ"
        Cells(1, 3).Value = "ФОТО"
        Cells(1, 4).Value = "ВИД ОБУВИ"
        Cells(1, 5).Value = "МАТЕРИАЛ ВЕРХА"
        Cells(1, 6).Value = "модель"
        Cells(1, 7).Value = "новый артикул"
        Cells(1, 8).Value = "код ТНВЭД"
        .Name = file_name
    End With
        
    With ThisWorkbook.Worksheets(1).Range("A:H")
        .Font.Bold = True
        .Font.Color = 0
        .Font.Size = 9
        .Font.Name = "Calibri"
        .EntireColumn.AutoFit
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .ColumnWidth = 25
        .RowHeight = 85
    End With

    Range("A1:E1").Interior.Color = 15917529
    Range("F1:H1").Interior.Color = 13431551
End Sub


