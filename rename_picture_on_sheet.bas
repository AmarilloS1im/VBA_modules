Attribute VB_Name = "rename_picture_on_sheet"
Sub rename_picture_on_sheet()
    Dim pic_name As String
    If TypeName(Selection) = "Picture" Then
        pic_name = Application.InputBox("������� ��� ������� ������ ��������� �������� ��� �������� ������ � ������: ", Type:=2)
        Selection.name = pic_name
    Else
        MsgBox "������� �������� �����������"
        Exit Sub
    End If
End Sub
