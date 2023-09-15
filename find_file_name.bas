Attribute VB_Name = "find_file_name"
Sub find_file_name()
    Dim array_file() As String
    Dim i As Integer
    Dim file_name As String
    
    array_file = Split(ThisWorkbook.Name, ".")
    
    For i = 0 To UBound(array_file) - 1
        file_name = file_name + array_file(i) + "."
    Next i
    file_name = Left(file_name, (Len(file_name) - 1))
    MsgBox file_name
End Sub
