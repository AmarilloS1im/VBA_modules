Attribute VB_Name = "rename_img_files"
Function find_extension(ByVal file_name As String)

    Dim splited_array() As String
    Dim extension_string As String
    
    splited_array = Split(file_name, ".")
    
    extension_string = extension_string & "." & CStr(splited_array(UBound(splited_array)))
    
    find_extension = extension_string
    
End Function
Function find_file_name(ByVal file_name As String)
    Dim array_file() As String
    Dim i As Integer
    Dim file_name_no_ext As String
    
    array_file = Split(file_name, ".")
    
    For i = 0 To UBound(array_file) - 1
        file_name_no_ext = file_name_no_ext + array_file(i) + "."
    Next i
    file_name_no_ext = Left(file_name_no_ext, (Len(file_name_no_ext) - 1))
    find_file_name = file_name_no_ext
End Function
Sub Rename_File()
    Dim fso As Object
    Dim myFolder As Object
    Dim myPath As String
    Dim newPath As String
    Dim myFile, myFiles(), i
    Dim new_collection As Collection
    Dim article As Range
    Dim file_name_without_ext As String
    Dim error_string As String
    Dim myDict As Object
    Dim extension
    
    
    
    
    
    
    myPath = Application.InputBox("Выберите путь к файлам", Type:=2)
    
    newPath = Application.InputBox("Выберите путь к папке, в которую будут скопированы файлы", Type:=2)
    
    If myPath = newPath Then
        MsgBox "Нельзя копировать файлы в оригинальный каталог!!!"
        Exit Sub
    End If
    
    
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Set myFolder = fso.GetFolder(myPath)
    
    If myFolder.Files.Count = 0 Then
        MsgBox "В папке «" & myPath & "» файлов нет"
        Exit Sub
    End If
    Set myDict = CreateObject("Scripting.Dictionary")
    
    On Error Resume Next
    For Each article In ActiveWorkbook.Worksheets(1).Range(Cells(1, 1), Cells(1, 1).End(xlDown))
        myDict.Add CStr(article), CStr(article.Offset(0, 6))
    Next article
    On Error GoTo 0
    
    error_string = ""
    'Загружаем в массив полные имена файлов
    ReDim myFiles(1 To myFolder.Files.Count)
    
    For Each myFile In myFolder.Files
        If myFile.Name <> "Thumbs.db" Then
            file_name_without_ext = find_file_name(myFile.Name)
            extension = find_extension(myFile.Name)
            
            If myDict.Exists(file_name_without_ext) Then
                FileCopy myPath & "\" & myFile.Name, newPath & "\" & myFile.Name
                If Dir(newPath & "\" & myDict(file_name_without_ext) & extension) = "" Then
                    Name newPath & "\" & myFile.Name As newPath & "\" & myDict(file_name_without_ext) & extension
                Else
                    Kill (newPath & "\" & myDict(file_name_without_ext) & extension)
                    Name newPath & "\" & myFile.Name As newPath & "\" & myDict(file_name_without_ext) & extension
                End If
            Else
                error_string = error_string & file_name_without_ext & extension & vbCrLf
            End If
        End If
    Next myFile

    
    If error_string <> "" Then
        MsgBox "Следующие файлы не найдены среди артикулов: " & vbCrLf & error_string
    Else
        MsgBox "Все файлы успешно переименованы"
    End If
    
End Sub
