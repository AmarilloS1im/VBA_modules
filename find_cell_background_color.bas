Attribute VB_Name = "find_cell_background_color"
Sub find_cell_background_color()
    With Application.ActiveWindow.ActiveCell
    MsgBox .Interior.Color
    End With
End Sub
