Attribute VB_Name = "Module1"
Sub filter()
    Dim filter_range As Range
    Dim filter_column As Integer
    Dim user_choice As Variant
    Dim user_choice_collection As New Collection
    Dim user_criterial As String
    Dim all_list_range
    Dim fill_range As Variant
    Dim coord_row
    Dim coord_column
    Dim item As Range
    Dim j As Variant
    Dim q As Variant
    Dim header_array()
    Dim n As Variant
    
    
    
    
    
    MsgBox "Откойте файл заказа в нужной папке" & vbCrLf & "*****************************************"
    'Открываем рабочий файл
    order_file = Application.GetOpenFilename("Excel files(*.xls*),*.xls*", 1, "Выбрать Excel файлы", , False)
    
    If VarType(order_file) = vbBoolean Then
        'Была нажата кнопка отмены-выход из процедуры
        Exit Sub
    End If
    
    'Выбираем строку к которой будут применены фильтры
    Set filter_range = Application.InputBox _
    ("Введите диапазон строки для фильтра" & vbCrLf & "например A2:AX2 или 2:2" _
    & vbCrLf & "Или выберете диапазон мышкой", Type:=8)
    
    'Выбираем номер колонки в котором будем выбирать условия фильтрации
    filter_column = Application.InputBox _
    ("Введите номер колонки где будете выбирать фильтр" & vbCrLf & "например 2", Type:=1)
    
    'Добавляем данные из колонки, которую выбрали ранее, и добавляем в коллекию уникальные значения.
    On Error Resume Next
    For Each user_choice In Range(Cells(filter_column + 1, filter_column), Cells(Rows.Count, filter_column).End(xlUp))
        user_choice_collection.Add user_choice.Value, user_choice.Value
    Next user_choice
    On Error GoTo 0
    
    'Обходим коллекцию в цикле и добавляем занчения в UserForms
    For Each user_choice In user_choice_collection
        UserForm1.ComboBox1.AddItem user_choice
    Next user_choice
    
    'Показываем пользователю UserForms
    UserForm1.Show
    'Присваиваем в переменную данные из UserForms которые выбрал пользователь в качестве фильтра
    user_criterial = UserForm1.ComboBox1.Value
    
    'Фильтруем документ по указанным ранее критериям
    With ActiveWorkbook.Worksheets(1).Range(filter_range.Address)
        .AutoFilter Field:=filter_column, Criteria1:=user_criterial, VisibleDropDown:=True
    End With
    
   Set all_list_range = ActiveWorkbook.Worksheets(1).Range(Cells(Rows.Count, 1).End(xlUp), Cells(1, Columns.Count))
   header_array = Array("артикул пп", "фото", "пол", "цвет поставщика", _
   "24", "25", "26", "27", "28", "29", "30", "31", "32", "33", "34", "35", "36", "37", "38", "39", "40", "41", "42" _
   , "43", "44", "45", "46", "47", "48")
   j = 6
   q = 1
   For Each n In header_array()
        coord_row = all_list_range.Find(n).Row + 1
        coord_column = all_list_range.Find(n).Column
        fill_range = ThisWorkbook.Worksheets(1) _
        .Range(Cells(coord_row, coord_column).Address(0, 0), _
        Cells(Rows.Count, coord_column).End(xlUp).Address(0, 0)).SpecialCells(xlCellTypeVisible).Address
        For Each item In Application.ActiveWorkbook.Worksheets(1).Range(fill_range)
            Application.Workbooks("тест1.xlsx").Worksheets(1).Cells(j, q) = item.Value
            j = j + 1
        Next item
        j = 6
        q = q + 1
    Next n
   
   
   
   
   
   
   
   
   
   
  
    'Очищаем UserForms
    UserForm1.ComboBox1.Clear
    
    
    
End Sub
