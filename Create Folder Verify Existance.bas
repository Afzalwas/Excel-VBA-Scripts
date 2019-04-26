Sub CreateFolderVerifyExistance()
FolderDestination = Range("B2").Value
FolderName = Range("B3").Value
Set fs = CreateObject("Scripting.FileSystemObject")
If fs.FolderExists(FolderDestination & FolderName) = True Then
End
ElseIf fs.FolderExists(FolderDestination & FolderName) = False Then
MkDir FolderDestination & FolderName
End If
End Sub

