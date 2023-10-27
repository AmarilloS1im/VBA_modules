Attribute VB_Name = "rename_picture_on_sheet"
Sub rename_picture_on_sheet()
    Dim pic_name As String
    If TypeName(Selection) = "Picture" Then
        pic_name = Application.InputBox("¬ведите им€ которое хотите присвоить картинке или выберите €чейку с именем: ", Type:=2)
        Selection.name = pic_name
    Else
        MsgBox "—начала выберите изображение"
        Exit Sub
    End If
End Sub
