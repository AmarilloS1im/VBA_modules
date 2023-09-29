Attribute VB_Name = "add_pictures_to_cells"
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

Sub add_picture_to_cell()
    Dim fso As Object
    Dim myFolder As Object
    Dim myPath As String
    Dim myFile, myFiles(), i
    Dim article As Range
    Dim file_name_without_ext As String
    Dim error_string As String
    Dim myDict As Object
    Dim extension As String
    Dim user_range As Range
    Dim cell_to_add_pic As Variant
    
    
    myPath = Application.InputBox("Выберите путь к файлам", Type:=2)
    
    Set user_range = Application.InputBox("Выберите диапазон артикулов", Type:=8)
    
    cell_to_add_pic = Application.InputBox("Укажите номер столбца куда добавить картинки. Счет ночинается отосительно столбка с артикулами", Type:=1)
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Set myFolder = fso.GetFolder(myPath)
    
    If myFolder.Files.Count = 0 Then
        MsgBox "В папке «" & myPath & "» файлов нет"
        Exit Sub
    End If
    
    Set myDict = CreateObject("Scripting.Dictionary")
    
    
    If user_range.Count = "1048576" Then
        Set user_range = ActiveWorkbook.Worksheets(1).Range(Cells(user_range.Row, user_range.Column), Cells(user_range.Row, user_range.Column).End(xlDown))
    Else

    End If
    
   
    'Загружаем в массив полные имена файлов
    ReDim myFiles(1 To myFolder.Files.Count)
    On Error Resume Next
    For Each myFile In myFolder.Files
        If myFile.Name <> "Thumbs.db" Then
            file_name_without_ext = find_file_name(myFile.Name)
            extension = find_extension(myFile.Name)
            myDict.Add CStr(file_name_without_ext), CStr(extension)
        Else
        
        End If
    Next myFile
    
    On Error GoTo 0
    i = 1
    For Each article In ActiveWorkbook.Worksheets(1).Range(user_range.Address)
        article.Offset(0, cell_to_add_pic).RowHeight = 110
        article.Offset(0, cell_to_add_pic).ColumnWidth = 28
        If myDict.Exists(article.Value) Then
            ActiveSheet.Pictures.Insert(myPath & "\" & article & myDict.Item(article.Value)).Select
            Selection.Cut
            article.Offset(0, cell_to_add_pic).Select
            ActiveSheet.Paste
            With ActiveSheet.Shapes(ActiveSheet.Shapes(i).Name)
               Set c = .TopLeftCell
              .LockAspectRatio = msoFalse
              .Left = c.Left
              .Top = c.Top
              .Width = c.Width
              .Height = c.Height
            End With
            i = i + 1
        Else
            error_string = error_string & article & vbCrLf
        End If
     Next article

    If error_string <> "" Then
        MsgBox "Следующие артикулы не найдены среди файлов в выбранной папке: " & vbCrLf & error_string
    Else
        MsgBox "Все картинки из указанной папки успешно добавлены"
    End If
    
    ActiveSheet.Shapes.SelectAll
    Selection.Placement = xlMoveAndSize
End Sub
