Attribute VB_Name = "Module1"
Function cross(ByRef column_name As String, ByRef row_name As String)
   Dim find_coor_y As Range
   Dim fing_coor_x As Range
   Dim final_coor As Range
   Set find_coor_x = Find(column_name, LookAt:=1)
   Set find_coor_y = Find(row_name, LookAt:=1)
   
   Set final_coor = Range(Cells(find_coor_x.Row, find_coor_x.column), Cells(find_coor_y.Row, find_coor_y.column))
   MsgBox final_coor.Address
   
    
    
End Function
