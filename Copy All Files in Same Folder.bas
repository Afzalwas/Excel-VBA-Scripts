Sub Copy_All_File_In_Same_Folder()
Set fs = CreateObject("Scripting.FileSystemObject")
oldpath = "c:\test"
newpath = "c:\test"
Prefix = "1_"
Set f = fs.GetFolder(oldpath)
Set NFile = f.Files
For Each pf1 In NFile
  NameFile = pf1.Name
  FileCopy oldpath & "\" & NameFile, newpath & "\" & Prefix & NameFile
Next
End Sub
