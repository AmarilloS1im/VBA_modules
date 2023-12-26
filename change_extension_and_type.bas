Attribute VB_Name = "change_extention_and_type"
Sub SaveAs_Mass()
    Dim sFolder As String, sFiles As String, sNonEx As String, sNewEx As String
    Dim wb As Workbook
    Dim lPos As Long, lFileFormat As Long, IsDelOriginal As Boolean
  
    'указываем новый формат файлов
    sNewEx = InputBox("Укажите новое расширение для файлов:", "Формат", "xlsx")
    'определяем числовой код формата файлов
    Select Case sNewEx
        Case "xlt": lFileFormat = 17
        Case "xla": lFileFormat = 18
        Case "xlsb": lFileFormat = 50
        Case "xlsx": lFileFormat = 51
        Case "xlsm": lFileFormat = 52
        Case "xltm": lFileFormat = 53
        Case "xltx": lFileFormat = 54
        Case "xlam": lFileFormat = 55
        Case "xls": lFileFormat = 56
        Case "csv": lFileFormat = 6
        'если указанный формат не соответсвует ни одному из существующих
        Case Else
            MsgBox "Формат '" & sNewEx & "' не поддерживается", vbCritical, "www.excel-vba.ru"
            Exit Sub
    End Select
  
    '   если надо просматривать файлы в той же папке, что и файл с кодом:
    '       sFolder = ThisWorkbook.Path
    'диалог запроса выбора папки с файлами
    With Application.FileDialog(msoFileDialogFolderPicker)
        If .Show = False Then Exit Sub
        sFolder = .SelectedItems(1)
    End With
    sFolder = sFolder & IIf(Right(sFolder, 1) = Application.PathSeparator, "", Application.PathSeparator)
    'запрашиваем - удалять ли исходные файлы после сохранения в новом формате
    IsDelOriginal = MsgBox("Удалять исходные файлы после пересохранения?", vbQuestion + vbYesNo, "") = vbYes
    'отключаем обновление экрана и показ системных сообщений
    Application.ScreenUpdating = 0
    Application.DisplayAlerts = 0
    Dim sh As Worksheet
    'просматриваем все файлы Excel в выбранной папке
    sFiles = Dir(sFolder & "*.xls*")
    Do While sFiles <> ""
        If sFiles <> ThisWorkbook.Name Then
            'получаем имя файла без расширения
            lPos = InStrRev(sFiles, ".")
            sNonEx = Mid(sFiles, 1, lPos)
            'открываем книгу
            Set wb = Application.Workbooks.Open(sFolder & sFiles, False)
            'сохраняем в новом формате и закрываем
            Select Case lFileFormat
            Case 24
                wb.Activate
                For Each sh In wb.Worksheets
                    sh.Select
                    wb.SaveAs sFolder & sNonEx & sh.Name & "." & sNewEx, lFileFormat
                Next
            Case Else
                wb.SaveAs sFolder & sNonEx & sNewEx, lFileFormat
            End Select
            wb.Close 0
            DoEvents
            'если надо удалить исходные файлы после преобразования
            If IsDelOriginal Then
                On Error Resume Next
                Kill sFolder & sFiles
                DoEvents
                On Error GoTo 0
            End If
        End If
        sFiles = Dir
    Loop
    'возвращаем обновление экрана и показ системных сообщений
    Application.ScreenUpdating = 1
    Application.DisplayAlerts = 1
    MsgBox "Файлы преобразованы", vbInformation, ""
End Sub
