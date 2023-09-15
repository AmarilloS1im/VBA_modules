Attribute VB_Name = "find_color_index"
Sub find_color_index()
    With Application.ActiveWindow.ActiveCell
    MsgBox .Font.Color
    End With
End Sub
