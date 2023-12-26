Attribute VB_Name = "packing_split"
Sub packing_split()
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
    
    
    last_col_on_sheet = Cells(1, Columns.count).End(xlToRight).Column
    
    Set user_range = Application.InputBox("Выберите диапазон с номерами коробок: ", Type:=8)
    
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
     'Application.ScreenUpdating = False
    
    first_carton = Range(Cells(user_range.Row, user_range.Column), Cells(user_range.Row, user_range.Column)).Value
    first_carton = CInt(Split(first_carton, "-")(0))
    
    last_carton = Range(Cells(user_range.Row, user_range.Column), Cells(user_range.Row, user_range.Column)).End(xlDown).Value

    If InStr(last_carton, "-") Then
        last_carton = last_carton = CInt(Split(last_carton, "-")(1))
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
            added_rows_count = added_rows_count + 1
            art_pos_count = art_pos_count + 1
            n = n + 1
        Next
        Selection.Delete Shift:=xlUp
        Set user_range = Range(Cells(user_range.Row + art_pos_count, user_range.Column), Cells(user_range.Row + art_pos_count, user_range.Column))
    Next
    Application.ScreenUpdating = True
    
End Sub
