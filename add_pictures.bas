Attribute VB_Name = "Module1"
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
    Dim oShp As Shape
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
    Dim addres_to_insert_cell
    
    
    
    
    
    
    myPath = Application.InputBox("Выберите путь к файлам", Type:=2)
    

    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Set myFolder = fso.GetFolder(myPath)
    
    If myFolder.Files.Count = 0 Then
        MsgBox "В папке «" & myPath & "» файлов нет"
        Exit Sub
    End If
    Set myDict = CreateObject("Scripting.Dictionary")
    
    On Error Resume Next
    For Each article In ThisWorkbook.Worksheets(1).Range(Cells(1, 1), Cells(1, 1).End(xlDown))
        article.Offset(0, 1).RowHeight = 60
        article.Offset(0, 1).ColumnWidth = 17
        myDict.Add CStr(article), CStr(article)
    Next article
    On Error GoTo 0
    
    'Загружаем в массив полные имена файлов
    ReDim myFiles(1 To myFolder.Files.Count)
    i = 1
    For Each myFile In myFolder.Files
        If myFile.Name <> "Thumbs.db" Then
            file_name_without_ext = find_file_name(myFile.Name)
            extension = find_extension(myFile.Name)
            
            If myDict.Exists(file_name_without_ext) Then
                ActiveSheet.Pictures.Insert(myPath & "\" & file_name_without_ext & extension).Select
                Selection.Cut
                Range("A1:A15").Find(file_name_without_ext).Offset(0, 1).Select
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
                MsgBox "нет картинки"
                
            End If
        End If
    Next myFile

    
    MsgBox "Картинки добавлены"
    
    ActiveSheet.Shapes.SelectAll
    Selection.Placement = xlMoveAndSize
End Sub
