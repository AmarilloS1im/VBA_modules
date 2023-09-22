Attribute VB_Name = "find_RU_codes_by_colors"
Dim user_range As String
Private Sub Workbook_Open()
    Application.OnKey "^+{q}", "find_RU_codes_by_colors.find_RU_codes_by_colors"
End Sub
Sub find_RU_codes_by_colors()
Attribute find_RU_codes_by_colors.VB_ProcData.VB_Invoke_Func = " \n14"
    Dim user_color() As String
    Dim my_range As range
    Dim output_string As String
    Dim i_color As Variant
    Dim cellcheck As range
    Dim ru_collection As New Collection
    Dim item As range
    
    If user_range = "" Then
        user_range = Application.InputBox("�������� �������� ������", Type:=8).Address
    End If
    
    
    user_color = Split(Application.InputBox("������� ����� ���� ��� ��������� ����� ������", Type:=2), " ")
    

    Set my_range = ThisWorkbook.Worksheets(1).range(user_range)
    For Each i_color In user_color
        On Error Resume Next
        For Each cellcheck In my_range
            ru_collection.Add cellcheck.Find(i_color, LookAt:=1).End(xlUp), _
            CStr(cellcheck.Find(i_color).End(xlUp))
        Next cellcheck
        On Error GoTo 0
    Next i_color
    
    
    For Each item In ru_collection
        output_string = output_string + item & " " & vbCrLf
    Next item
    If output_string = "" Then
        MsgBox "������ �� ������� � ��������� ���������"
    Else
        MsgBox output_string
    End If
    
End Sub
