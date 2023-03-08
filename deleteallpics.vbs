Sub DeleteAllPics()
Dim Pic As Object
For Each Pic In ActiveSheet.Pictures
Pic.Delete
Next Pic
End Sub