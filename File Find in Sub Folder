Sub FileFindInSubfolders()
   Dim StrFile As String, objFSO, destRow As Long
   Dim mainFolder, mySubFolder
   
   mFolder = Range("B2").Value ' folder
   fname = "*.xl*" ' file name

   Set objFSO = CreateObject("Scripting.FileSystemObject")
   FolderDestination = Sheets(1).Range("B2").Value
   FolderName = Sheets(1).Range("B3").Value
   mFolder = FolderDestination & FolderName & "\"
   
   Set mainFolder = objFSO.GetFolder(mFolder)
   
   StrFile = Dir(mFolder & "\" & fname)
   If StrFile <> "" Then
       MsgBox mFolder & "\" & StrFile & " found"
   
   
   Else
     For Each mySubFolder In mainFolder.SubFolders
       StrFile = Dir(mySubFolder & "\" & fname)
       If StrFile <> "" Then
          MsgBox mySubFolder & "\" & StrFile & " found"
          
           
       End If
     Next
   End If
End Sub


