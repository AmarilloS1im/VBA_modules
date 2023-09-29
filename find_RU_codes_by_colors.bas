Attribute VB_Name = "find_RU_codes_by_colors"

Sub find_RU_codes_by_colors()
Attribute find_RU_codes_by_colors.VB_ProcData.VB_Invoke_Func = " \n14"
    Dim user_color() As String
    Dim my_range As Range
    Dim output_string As String
    Dim i_color As Variant
    Dim cellcheck As Range
    Dim i As Variant
    Dim myDict As Object
    Dim error_string As String
    Dim user_range As String
    Set myDict = CreateObject("Scripting.Dictionary")
    
    
    user_range = "A4:CF50"
  
    'user_range = Application.InputBox("выберите диапазон поиска", Type:=8).Address
    'Специально захардкодил рейндж, чтоб не выбирать постоянно

    
    
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
        output_string = output_string + i & "--" & myDict.Item(i) & " " & vbCrLf
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
    
End Sub
