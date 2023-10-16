Attribute VB_Name = "Module1"
Sub Save_Object_As_Picture_NamesFromCells()
    Dim li As Long, oObj As Shape, wsSh As Worksheet, wsTmpSh As Worksheet
    Dim sImagesPath As String, sName As String
    Dim lNamesCol As Long, s As String
 
    s = InputBox("������� ����� ������� � ������� ��� ��������" & vbNewLine & _
                 "(0 - ������� � ������� ���� ��������)")
    If StrPtr(s) = 0 Then Exit Sub
    lNamesCol = Val(s)
 
    sImagesPath = Application.InputBox("�������� ���� � ������: ", Type:=2) & "\"
    If Dir(sImagesPath, 16) = "" Then
        MkDir sImagesPath
    End If
'    On Error Resume Next
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Set wsSh = ActiveSheet
    Set wsTmpSh = ActiveWorkbook.Sheets.Add
    For Each oObj In wsSh.Shapes
        If oObj.Type = 13 Then
            oObj.Copy
            If lNamesCol = 0 Then
                sName = oObj.TopLeftCell.Value
            Else
                sName = wsSh.Cells(oObj.TopLeftCell.Row, lNamesCol).Value
            End If
            '���� � ������ ���� �������, �����������
            '��� ������������� � �������� ���� ��� ������ - �������
            sName = CheckName(sName)
            '���� sName � ���������� ����� - ���� ��� unnamed_ � ���������� �������
            If sName = "" Then
                li = li + 1
                sName = "unnamed_" & li
            End If
            With wsTmpSh.ChartObjects.Add(0, 0, oObj.Width, oObj.Height).Chart
                .ChartArea.Border.LineStyle = 0
                .Parent.Select
                .Paste
                .Export Filename:=sImagesPath & sName & ".jpg", FilterName:="JPG"
                .Parent.Delete
            End With
        End If
    Next oObj
    Set oObj = Nothing: Set wsSh = Nothing
    wsTmpSh.Delete
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox "������� ��������� � �����: " & sImagesPath, vbInformation
End Sub
'---------------------------------------------------------------------------------------
' Procedure : CheckName
' Purpose   : ������� �������� ������������ �����
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

