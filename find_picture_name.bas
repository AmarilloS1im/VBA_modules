Attribute VB_Name = "find_picture_name"
Sub find_picture_name()
    If TypeName(Selection) = "Picture" Then
        MsgBox Selection.Name
    Else
        MsgBox "Сначала выберите изображение"
        Exit Sub
    End If
End Sub
