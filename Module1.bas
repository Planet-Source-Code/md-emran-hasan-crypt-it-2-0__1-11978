Attribute VB_Name = "Module1"
Public fMainForm As frmMain


Sub Main()
    Set fMainForm = New frmMain
    fMainForm.Show
End Sub


Public Function Crypt(ByVal Text As String, ByVal Key As String) As String
   For i = 1 To Len(Text)
   A = i Mod Len(Key): If A = 0 Then A = Len(Key)
   Crypt = Crypt & Chr(Asc(Mid(Key, A, 1)) Xor Asc(Mid(Text, i, 1)))
Next i
End Function

