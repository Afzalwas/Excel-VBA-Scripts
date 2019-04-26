Sub Delete_Whole_Folder()
'Delete whole folder without removing the files first like in DeleteExample4
    Dim FSO As Object

    Set FSO = CreateObject("scripting.filesystemobject")

    With ThisWorkbook.Sheets(2)
        LR = .Cells(.Rows.Count, "P").End(xlUp).Row
        MyPath = .Range("P2:P" & LR)
    End With
    For i = LBound(MyPath) To UBound(MyPath)
    
    If Right(MyPath(i, 1), 1) = "\" Then
        MyPath = Left(MyPath(i, 1), Len(MyPath(i, 1)) - 1)
    End If

    If FSO.FolderExists(MyPath(i, 1)) = False Then
        MsgBox MyPath & " doesn't exist"
        Exit Sub
    End If

    FSO.deletefolder MyPath(i, 1)
Next i
End Sub
