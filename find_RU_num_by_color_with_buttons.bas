Attribute VB_Name = "Find_RU_by_color_with_buttons"
Public first_button_exist As String
Public second_button_exist As String
Public quit_button_exist As String
Public user_range As String
Function choose_range()
    user_range = Application.InputBox("выберите диапазон поиска", Type:=8).Address
    choose_range = user_range
End Function
Function terminate_sub()
    Dim shp As Shape
    For Each shp In ActiveSheet.Shapes
        shp.Delete
    Next
    first_button_exist = ""
    second_button_exist = ""
    quit_button_exist = ""
    user_range = ""
    ThisWorkbook.Application.EnableEvents = False
End Function
Function create_button(ByVal ra As Range, Optional ByVal ButtonColor As Long = 255, _
                       Optional ByVal ButtonName$ = "Запуск", Optional ByVal MacroName As String = "")
    ' Функция рисует автофигуру (прямоугольник) поверх диапазона ячеек ra
    ' и окрашивает созданную кнопку (с названием ) в цвет Button_color
    ' Созданной кнопке назначается макрос MacroName
    On Error Resume Next: Err.Clear
    w = ra.Width: h = ra.Height: l = ra.Left: t = ra.Top
    w = IIf(w >= 10, w, 50): h = IIf(h >= 10, h, 50)    ' не создаём маленькие кнопки - минимум 10*10

    ' добавляем кнопку на лист
    Dim sha As Shape: Set sha = ActiveSheet.Shapes.AddShape(msoShapeRoundedRectangle, l, t, w, h)
    With sha    ' оформляем автофигуру
        .Fill.Visible = msoTrue: .Fill.Solid
        .Fill.ForeColor.RGB = ButtonColor: .Fill.Transparency = 0.3
        .Fill.BackColor.RGB = vbWhite
        .Fill.OneColorGradient msoGradientHorizontal, 4, 0 ' градиентная заливка
        .Adjustments.item(1) = 0.23: .Placement = xlFreeFloating
        .OLEFormat.Object.PrintObject = False    ' кнопки не выводятся на печать
        .Line.Weight = 0.25: .Line.ForeColor.RGB = vbBlack ' делаем тонкий черный контур
        With .TextFrame    ' добавляем и форматируем текст
            .Characters.Text = ButtonName$ ' добавляем текст
            With .Characters.Font ' изменяем начертание текста
                .Size = IIf(h >= 16, 10, 8): .Bold = True:
                .Color = vbBlack: .Name = "Arial" ' цвет и шрифт
            End With
            .HorizontalAlignment = xlCenter: .VerticalAlignment = xlVAlignCenter
        End With
        .OnAction = MacroName    ' назначаем кнопке макрос (если он задан в параметрах)
    End With
End Function
Sub find_RU_num_by_color()
    Dim user_color() As String
    Dim my_range As Range
    Dim output_string As String
    Dim i_color As Variant
    Dim cellcheck As Range
    Dim i As Variant
    Dim myDict As Object
    Dim error_string As String
    'Dim user_range As String
    Set myDict = CreateObject("Scripting.Dictionary")
    
    Dim ra As Range
    Dim ra_2 As Range
    Dim q_b_range As Range
    Set ra = ActiveWorkbook.ActiveSheet.Range("A1:C3")
    Set ra_2 = ActiveWorkbook.ActiveSheet.Range("D1:E3")
    Set q_b_range = ActiveWorkbook.ActiveSheet.Range("J1:L3")
    
    If first_button_exist = "" Then
        Call create_button(ra, 3574037, "Найти цвет", "find_RU_num_by_color")
        first_button_exist = "buttno allready exist"
    Else
    End If
    
    If second_button_exist = "" Then
        Call create_button(ra_2, 14318848, "Задать/Сменить диапазон", "choose_range")
        second_button_exist = "buttno allready exist"
    Else
    End If
    
    If quit_button_exist = "" Then
        Call create_button(q_b_range, 5460991, "Закончить работу и удалить кнопки", "terminate_sub")
        quit_button_exist = "buttno allready exist"
    Else
    End If
    

    If user_range = "" Then
        MsgBox "Задайте диапазон поиска"
        Exit Sub
    Else
        user_color = Split(Application.InputBox("Введите цвета один или несколько через запятую", Type:=2), ",")
        
        Set my_range = ActiveWorkbook.Worksheets(1).Range(user_range)
        
        For Each i_color In user_color
            On Error Resume Next
            For Each cellcheck In my_range
                'Чтобы изменить поиск с точного совпадения на частичное совпадение нужно
                'поменять значение параметра LookAt с 1 на 2
                If Not cellcheck.Find(i_color, LookAt:=1) Is Nothing Then
                    myDict.Add Cells(my_range.Row, cellcheck.Column), Cells(cellcheck.Row, cellcheck.Column)
                 Else
                 End If
            Next cellcheck
            On Error GoTo 0
        Next i_color
        
        
        For Each i In myDict
            output_string = output_string + i & "--" & myDict.item(i) & " " & vbCrLf
        Next i
        
        
        For Each i_color In user_color
            With my_range
                Set c = .Find(i_color, LookIn:=xlValues)
                If Not c Is Nothing Then
                Else
                    error_string = error_string & i_color & vbCrLf
                End If
            End With
        Next i_color
        
        
        If output_string = "" Then
            MsgBox "Следующие цвета не найдены в указанном диапазоне:" & vbCrLf & error_string
        Else
            If error_string = "" Then
                MsgBox output_string
            Else
                MsgBox output_string & vbCrLf & vbCrLf & "Следующие цвета не найдены или указаны некорректно:" & vbCrLf & error_string
            End If
        End If
    End If
End Sub






