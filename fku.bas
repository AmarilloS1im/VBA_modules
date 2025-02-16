Attribute VB_Name = "fku"
Option Explicit

Sub find_price_list_duplicates()
    Dim i As Long
    Dim first_row As Long
    Dim last_row As Variant
    Dim list_range As String
    Dim rng As Range
    Dim row_i As Long
    Dim col_i As Long
    Dim start_row As Long
    Dim last_row_total_data As Variant
    Dim total_data_list_range As String
    Dim total_data_list_range_2 As String
    Dim total_del_rows As Long
    Dim last_col As Variant
    Dim curren_row As Variant
    Dim prev_row As Variant
    Dim current_row_price As Variant
    Dim prev_row_price As Variant
    Dim lists_num As Integer
    
    Dim tmp_dict As Object
    Set tmp_dict = CreateObject("Scripting.Dictionary")
    
    
    
    
    lists_num = Application.InputBox("Введите количество прайс-листов")
    lists_num = lists_num + 1
    MsgBox ("Обработка файла началась")
    first_row = 10
    Worksheets(1).Activate
    ActiveWorkbook.Sheets.Add
    Worksheets(1).Name = "Общие данные"
    start_row = 2
    For i = 2 To lists_num
        row_i = 10
        col_i = 2
        Worksheets(i).Activate
        last_row = Range("D" & Rows.Count).End(xlUp).Row
        list_range = "A10" & ":" & "D" & last_row
        For Each rng In ActiveWorkbook.ActiveSheet.Range(list_range)
'            Cells(row_i, col_i + 3).Value = Cells(row_i, col_i).Value & "_" & ActiveWorkbook.Sheets(i).Name & "_" & Cells(row_i, col_i + 2).Value
            ActiveWorkbook.Worksheets("Общие данные").Cells(1, 1) = "№"
            ActiveWorkbook.Worksheets("Общие данные").Cells(1, 2) = "Артикул"
            ActiveWorkbook.Worksheets("Общие данные").Cells(1, 3) = "Наименование"
            ActiveWorkbook.Worksheets("Общие данные").Cells(1, 4) = "Цена"
            ActiveWorkbook.Worksheets("Общие данные").Cells(1, 5) = "Сцепка"
            ActiveWorkbook.Worksheets("Общие данные").Cells(1, 6) = "Проверка"
            ActiveWorkbook.Worksheets("Общие данные").Cells(start_row, 1) = Cells(row_i, 1).Value
            ActiveWorkbook.Worksheets("Общие данные").Cells(start_row, 2) = Cells(row_i, 2).Value
            ActiveWorkbook.Worksheets("Общие данные").Cells(start_row, 3) = Cells(row_i, 3).Value
            ActiveWorkbook.Worksheets("Общие данные").Cells(start_row, 4) = Cells(row_i, 4).Value
            ActiveWorkbook.Worksheets("Общие данные").Cells(start_row, 5) = Cells(row_i, 5).Value


            row_i = row_i + 1
            start_row = start_row + 1
            If row_i = last_row + 1 Then
                Exit For
            End If
        Next
    Next i
    
    Worksheets("Общие данные").Activate
    last_row_total_data = Range("D" & Rows.Count).End(xlUp).Row
    total_data_list_range = "F2" & ":" & "F" & last_row_total_data
    
    Range("A1:F1").Select
    Selection.AutoFilter
    For Each rng In ActiveWorkbook.ActiveSheet.Range(total_data_list_range)
        rng.FormulaR1C1 = "=COUNTIF(C[-4],RC[-4])"
    Next
    Range("F2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

    row_i = 1
    total_del_rows = 1
    Do While total_del_rows > 0
    total_del_rows = 0
    last_col = Cells(1, Columns.Count).Column
    last_row_total_data = Range("D" & Rows.Count).End(xlUp).Row
    For i = 2 To last_row_total_data
        If Cells(i, 6).Value = 1 Or Cells(i, 6).Value = 0 Then
            ActiveWorkbook.ActiveSheet.Range(Cells(i, 1), Cells(i, last_col)).EntireRow.Delete
            total_del_rows = total_del_rows + 1
        End If
        row_i = row_i + 1
    Next
    Loop
    
    ActiveWorkbook.Worksheets("Общие данные").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Общие данные").AutoFilter.Sort.SortFields.Add2 Key _
        :=Range("B1"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("Общие данные").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    total_del_rows = 1
    Do While total_del_rows > 0
    total_del_rows = 0
    last_row_total_data = Range("D" & Rows.Count).End(xlUp).Row
    prev_row = ""
    prev_row_price = 0
    For i = 2 To last_row_total_data
        curren_row = UCase(CStr(Cells(i, 2).Value))
        current_row_price = Cells(i, 4).Value
        If curren_row = prev_row Then
            If current_row_price > prev_row_price Then
                ActiveWorkbook.ActiveSheet.Range(Cells(i, 1), Cells(i, last_col)).EntireRow.Delete
                total_del_rows = total_del_rows + 1
            Else
                ActiveWorkbook.ActiveSheet.Range(Cells(i - 1, 1), Cells(i - 1, last_col)).EntireRow.Delete
                total_del_rows = total_del_rows + 1
            End If
        End If
        prev_row = curren_row
        prev_row_price = current_row_price
        row_i = row_i + 1
    Next
    Loop
    
    last_row_total_data = Range("D" & Rows.Count).End(xlUp).Row
    For i = 2 To last_row_total_data
        tmp_dict.Add UCase(CStr(Cells(i, 2).Value)), Cells(i, 4).Value
        row_i = row_i + 1
    Next

    For i = 2 To lists_num
      Worksheets(i).Activate
      total_del_rows = 1
      Do While total_del_rows > 0
        total_del_rows = 0
        last_row = Range("D" & Rows.Count).End(xlUp).Row
        list_range = "A10" & ":" & "D" & last_row
        row_i = 10
        col_i = 2
        For Each rng In ActiveWorkbook.ActiveSheet.Range(list_range)
            If tmp_dict.Exists(UCase(CStr(Cells(row_i, col_i).Value))) Then
                If tmp_dict(UCase(CStr(Cells(row_i, col_i).Value))) <> Cells(row_i, col_i + 2).Value Then
                    ActiveWorkbook.ActiveSheet.Range(Cells(row_i, 1), Cells(row_i, last_col)).EntireRow.Delete
                    total_del_rows = total_del_rows + 1
                End If
            End If
            row_i = row_i + 1
        Next
       Loop
    Next
    MsgBox ("Обработка завершена")
End Sub
