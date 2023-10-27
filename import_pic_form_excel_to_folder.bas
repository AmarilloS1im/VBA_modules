Attribute VB_Name = "import_pic_form_excel_to_folder"
Sub Import_pictures_form_excel_to_folder()
    Dim li As Long, oObj As Shape, wsSh As Worksheet, wsTmpSh As Worksheet
    Dim sImagesPath As String, sName As String
    Dim lNamesCol As Long, s As String
    Dim picture_type
 
    s = InputBox("Укажите номер столбца с именами для картинок" & vbNewLine & _
                 "(0 - столбец в котором сама картинка)")
    If StrPtr(s) = 0 Then Exit Sub
    lNamesCol = Val(s)
 
    sImagesPath = Application.InputBox("Выберите путь к папке в которую нужно сохранить картинки: ", Type:=2) & "\"
    
    picture_type = Application.InputBox("Выберите тип изображения:" & vbCrLf & "1 - Автофигуры" _
    & vbCrLf & "3 - Диаграммы " & vbCrLf & "11 - Связанное изображение" _
    & vbCrLf & "13 - Картинки" & vbCrLf & "Если не получилось сохранить картинки с первого раза, попробуйте поменять типы изображения" _
    , Type:=1)
    
    '& vbCrLf & ""https://learn.microsoft.com/ru-ru/office/vba/api/Office.MsoShapeType"" & vbCrLf & "Если не получилось сохранить картинки с первого раза, попробуйте поменять типы изображения"
    
    
    If Dir(sImagesPath, 16) = "" Then
        MkDir sImagesPath
    End If
'    On Error Resume Next
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Set wsSh = ActiveSheet
    Set wsTmpSh = ActiveWorkbook.Sheets.Add
    For Each oObj In wsSh.Shapes
        If oObj.Type = picture_type Then
            oObj.Copy
            If lNamesCol = 0 Then
                sName = oObj.TopLeftCell.Value
            Else
                sName = wsSh.Cells(oObj.TopLeftCell.Row, lNamesCol).Value
            End If
            'если в ячейке были символы, запрещенные
            'для использования в качестве имен для файлов - удаляем
            sName = CheckName(sName)
            'если sName в результате пусто - даем имя unnamed_ с порядковым номером
            If sName = "" Then
                li = li + 1
                sName = "unnamed_" & li
            End If
            With wsTmpSh.ChartObjects.Add(0, 0, oObj.Width, oObj.Height).Chart
                .ChartArea.Border.LineStyle = 0
                .Parent.Select
                .Paste
                .Export Filename:=sImagesPath & sName & ".jpeg", FilterName:="JPEG"
                .Parent.Delete
            End With
        End If
    Next oObj
    Set oObj = Nothing: Set wsSh = Nothing
    wsTmpSh.Delete
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox "Объекты сохранены в папке: " & sImagesPath, vbInformation
End Sub
'---------------------------------------------------------------------------------------
' Procedure : CheckName
' Purpose   : Функция проверки правильности имени
'---------------------------------------------------------------------------------------
Function CheckName(sName As String)
    Dim objRegExp As Object
    Dim s As String
    Set objRegExp = CreateObject("VBScript.RegExp")
    objRegExp.Global = True: objRegExp.IgnoreCase = True
    objRegExp.Pattern = "[:,\\,/,?,\*,\<,\>,\',\|,""""]"
    s = objRegExp.Replace(sName, "")
    CheckName = s
End Function

