VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   3036
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4584
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False





Private Sub OptionButton1_Click()
    If OptionButton1.Value = True Then
        flag = False
        Unload Me
    End If

End Sub

Private Sub OptionButton2_Click()
        If OptionButton2.Value = True Then
            flag = True
            Unload Me
        End If

End Sub

Private Sub UserForm_Click()

End Sub
