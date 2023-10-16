Attribute VB_Name = "merge_cells"
Sub merge_cells()
    Dim cellcheck As Range
    Dim prev_cell As String
    Dim curr_cell As String
    Dim user_range As Range
    Dim start_range As Range
    Dim end_range As Range
    Dim count_num As Long
    Dim user_merge_range As Range
    
    
    Set user_range = Application.InputBox("Выберите диапазон с артикулами: ", Type:=8)
    Set user_merge_range = Application.InputBox("Выберите диапазон в котором хотите объеденить ячейки: ", Type:=8)
    
    
    Set user_range = Range(Cells(user_range.Row, user_range.Column), Cells(user_range.Count, user_range.Column).End(xlUp))
    Set start_range = Range(Cells(user_range.End(xlUp).Row, user_range.End(xlUp).Column), _
    Cells(user_range.End(xlUp).Row, user_range.End(xlUp).Column))
    
    Application.ScreenUpdating = False

    count_num = 0
    prev_cell = start_range.Cells
    For Each cellcheck In user_range
        curr_cell = cellcheck.Cells
        If prev_cell = curr_cell Then
            count_num = count_num + 1
            Set end_range = Range(Cells(cellcheck.Row, cellcheck.Column), Cells(cellcheck.Row, cellcheck.Column))
        Else
            Range(Cells(start_range.Row, user_merge_range.Column), Cells(end_range.Row, user_merge_range.Column)).Select
            Selection.Merge
            Set start_range = Range(Cells(end_range.Row + 1, start_range.Column), Cells((end_range.Row + 1), start_range.Column))
        prev_cell = curr_cell
        count_num = 0
        End If
    Next cellcheck
    Range(Cells(start_range.Row, user_merge_range.Column), Cells(end_range.Row, user_merge_range.Column)).Select
    Selection.Merge
    
    Application.ScreenUpdating = True
End Sub

