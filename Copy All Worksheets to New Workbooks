Public Sub CopySheetToNewWorkbook()
    With Application
        .DisplayAlerts = False
        .ScreenUpdating = False
    End With
            On Error Resume Next
            ChDrive Left(ActiveWorkbook.Path, 1)
            Folder_Path = CStr(ActiveWorkbook.Path) & "/"
            ChDir Folder_Path
            On Error GoTo 0

Set wbBook = ActiveWorkbook
Set wsActive = wbBook.Sheets("Master")

     For Each wsSheet In wbBook.Worksheets
        If wsSheet.Name <> wsActive.Name Then
            wsSheet.Activate
            fname = wsSheet.Name '& ".xls"
                ActiveSheet.Copy
                ActiveWorkbook.SaveAs Filename:=fname, FileFormat:=xlExcel8
                ActiveWorkbook.Close
        End If
    Next wsSheet
    
    wsActive.Activate
    
                With Application
                    .DisplayAlerts = True
                    .ScreenUpdating = True
                End With
End Sub
