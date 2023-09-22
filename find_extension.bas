Attribute VB_Name = "find_extension"
Sub find_extension(ByVal file_name As String)

    Dim splited_array() As String
    Dim extension_string As String
    
    splited_array = Split(file_name, ".")
    
    extension_string = extension_string & "." & CStr(splited_array(UBound(splited_array)))
    MsgBox extension_string
End Sub


