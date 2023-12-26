Attribute VB_Name = "pk_split_sort_size_add"
Sub packing_split_and_sort_and_size_add()
    Dim user_range As Range
    Dim article As Range
    Dim size As Range
    Dim size_range As Range
    Dim offset_count As Integer
    Dim last_col_on_sheet As Long
    Dim added_rows_count As Integer
    Dim art_pos_count As Long
    Dim i As Long
    Dim j As Long
    Dim count_len As Long
    Dim first_carton As Variant
    Dim last_carton As Variant
    Dim for_index As Long
    Dim n As Variant
    Dim cartons_count As Integer
    Dim carton_list As New Collection
    Dim control_num As Variant
    Dim loop_count As Long
    Dim size_range_offset_count As Long
    Dim new_range As Range
    Dim count_new_rows As Long
    
    
    MsgBox ("ВНИМАНИЕ!" _
    & vbCrLf & "Колонка с номерами коробок, количеством коробок, количеством в коробке и общее количество в коробке должны стоять левее размерной сетки")
    
    last_col_on_sheet = Cells(1, Columns.Count).End(xlToRight).Column
    
    Set user_range = Application.InputBox("Выберите диапазон с номерами коробок: ", Type:=8)
    
    Set size_range = Application.InputBox("Выберите диапазон с размерной сеткой: ", Type:=8)
    
    cell_to_add_size = Application.InputBox("Укажите номер столбца куда добавить размеры." _
    & vbCrLf & "Счет с первого столбца.", Type:=1)
    
    cell_to_add_quantity = Application.InputBox("Укажите номер столбца куда добавить количество." _
    & vbCrLf & "Счет с первого столбца.", Type:=1)
    
    Application.ScreenUpdating = False
    
    'Считаем длину изначального рейнджа с артикулами
    For Each article In Range(user_range.Address)
        count_len = count_len + 1
    Next article
    
    'Сортируем коробки по возрастанию
    Do
        loop_count = 0
        control_num = Range(Cells(user_range.Row, user_range.Column), Cells(user_range.Row, user_range.Column)).Value
        If InStr(control_num, "-") Then
            control_num = CInt(Split(control_num, "-")(0))
        Else
        End If

        For Each carton_num In Range(user_range.Address)
            
            If InStr(carton_num, "-") Then
                first_num = CInt(Split(carton_num, "-")(0))
                last_num = CInt(Split(carton_num, "-")(1))
            Else
                first_num = CInt(carton_num)
                last_num = CInt(carton_num)
            End If
        
            If first_num >= control_num Then
                control_num = last_num
            Else
                Rows(carton_num.Row - 1).Select
                Selection.Cut
                Rows(carton_num.Row + 1).Select
                Selection.Insert Shift:=xlDown
                control_num = last_num
                loop_count = loop_count + 1
            End If
        Next carton_num
     Loop While loop_count <> 0
     
    'Сплитуем коробки
    first_carton = Range(Cells(user_range.Row, user_range.Column), Cells(user_range.Row, user_range.Column)).Value
    first_carton = CInt(Split(first_carton, "-")(0))
    
    last_carton = Range(Cells(user_range.Row, user_range.Column), Cells(user_range.Row, user_range.Column)).End(xlDown).Value

    If InStr(last_carton, "-") Then
        last_carton = CInt(Split(last_carton, "-")(1))
    Else
        last_carton = CInt(last_carton)
    End If
    
    For i = first_carton To last_carton
        carton_list.Add i
    Next

    n = 1
    For i = 1 To count_len
        user_range.Select
        Set article = Range(Cells(user_range.Row, user_range.Column), Cells(user_range.Row, user_range.Column))
        added_rows_count = 1
        art_pos_count = 0
        cartons_count = CInt(article.Offset(0, 1).Value)
        For j = 1 To cartons_count
            Range(Cells(article.Row + art_pos_count, 1), Cells(article.Row + art_pos_count, last_col_on_sheet)).Copy
            Range(Cells(article.Row + art_pos_count, 1), Cells(article.Row + art_pos_count, last_col_on_sheet)).Offset(1, 0).Select
            Selection.Insert Shift:=xlDown
            Cells(article.Row + art_pos_count, article.Column).Value = carton_list(n)
            count_new_rows = count_new_rows + 1
            
            added_rows_count = added_rows_count + 1
            art_pos_count = art_pos_count + 1
            n = n + 1
        Next
        Selection.Delete Shift:=xlUp
        Set user_range = Range(Cells(user_range.Row + art_pos_count, user_range.Column), Cells(user_range.Row + art_pos_count, user_range.Column))
    Next
    
    'Разбиваем строки по размерам и количествам
    Set new_range = Range(Cells(user_range.Row - count_new_rows, user_range.Column), Cells(user_range.Row - 1, user_range.Column))
    Set user_range = Range(new_range.Address)
    count_len = 0
    For Each article In Range(user_range.Address)
        count_len = count_len + 1
    Next article
    size_count = 0
    size_range_offset_count = 1
    For i = 1 To count_len
        user_range.Select
        Set article = Range(Cells(user_range.Row, user_range.Column), Cells(user_range.Row, user_range.Column))
        added_rows_count = 1
        art_pos_count = 0
        For Each size In size_range.Offset(size_range_offset_count, 0)
            If size.Value <> "" Then
                Cells(article.Row + art_pos_count, cell_to_add_quantity).Value = size.Value
                Cells(article.Row + art_pos_count, cell_to_add_size).Value = Cells(size_range.Row, size.Column).Value
                Range(Cells(article.Row + art_pos_count, 1), Cells(article.Row + art_pos_count, last_col_on_sheet)).Offset(1, 0).Select
                Selection.Insert Shift:=xlDown
                Cells(article.Row + art_pos_count + 1, article.Column).Value = article.Value
                added_rows_count = added_rows_count + 1
                art_pos_count = art_pos_count + 1
                size_range_offset_count = size_range_offset_count + 1
            Else
            End If
         Next size
         Selection.Delete Shift:=xlUp
         Set user_range = Range(Cells(user_range.Row + art_pos_count, user_range.Column), Cells(user_range.Row + art_pos_count, user_range.Column))
      Next
    
    Application.ScreenUpdating = True
    
End Sub
