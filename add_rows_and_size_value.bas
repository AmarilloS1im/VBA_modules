Attribute VB_Name = "add_rows_and_size_value"
Sub add_rows_and_size_value()
    Dim user_range As Range
    Dim article As Range
    Dim size As Range
    Dim size_range As Range
    Dim offset_count As Integer
    Dim last_col_on_sheet As Long
    Dim added_rows_count As Integer
    Dim art_pos_count As Long
    Dim size_range_offset_count As Long
    Dim i As Long
    Dim count_len As Long
    
    last_col_on_sheet = Cells(1, Columns.count).End(xlToRight).Column
    
    Set user_range = Application.InputBox("Выберите диапазон с артикулами: ", Type:=8)
    
    Set size_range = Application.InputBox("Выберите диапазон с размерной сеткой: ", Type:=8)
    
    cell_to_add_size = Application.InputBox("Укажите номер столбца куда добавить размеры." _
    & vbCrLf & "Счет с первого столбца.", Type:=1)
    
    cell_to_add_quantity = Application.InputBox("Укажите номер столбца куда добавить количество." _
    & vbCrLf & "Счет с первого столбца.", Type:=1)
    
    'Считаем длину изначального рейнджа с артикулами
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
End Sub
