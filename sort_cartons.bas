Attribute VB_Name = "sort_carton"
Sub sort_carton()
    Dim user_range As Range
    Dim article As Range
    Dim last_col_on_sheet As Long
    Dim count_len As Long
    Dim cartons_count As Integer
    Dim carton_num As Range
    Dim first_num As Long
    Dim last_num As Long
    Dim control_num As Variant
    Dim loop_count As Long
    
    
    last_col_on_sheet = Cells(1, Columns.Count).End(xlToRight).Column
    
    Set user_range = Application.InputBox("Выберите диапазон с номерами коробок: ", Type:=8)
    
    Application.ScreenUpdating = False
    
    'Считаем длину изначального рейнджа с артикулами
    For Each article In Range(user_range.Address)
        count_len = count_len + 1
    Next article
    
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
     Application.ScreenUpdating = True
End Sub
