Sub GeneratingXLSFilespathfromfolderandsubfolderssnb()
FolderDestination = Sheets(1).Range("B2").Value
FolderName = Sheets(1).Range("B3").Value
    c00 = FolderDestination & FolderName & "\*.xls"
    sn = Application.Transpose(Split(CreateObject("wscript.shell").exec("cmd /c Dir " _
    & """" & c00 & """" & " /b /s").StdOut.ReadAll, vbCrLf))
    Cells(2, 1).Resize(UBound(sn)) = sn
End Sub
