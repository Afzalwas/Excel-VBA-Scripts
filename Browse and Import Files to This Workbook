Sub BrowseSelectFile()

Dim Folder_Path As String

ActiveSheet.Unprotect "password"

    On Error Resume Next
    ChDrive Left(ActiveWorkbook.Path, 1)
    Folder_Path = CStr(ActiveWorkbook.Path) & "/"
    ChDir Folder_Path
    On Error GoTo 0
        
        With Application
            .DisplayAlerts = False
            .ScreenUpdating = False
        End With

            Set wbBook = ActiveWorkbook
            Set wsActive = wbBook.Sheets("Master")

                Dim fn
                fn = Application.GetOpenFilename("Excel Files (*.xls; *.xlsx),*.xls;*.xlsx")  'can add parameters. See help for details.
                    If fn <> False Then
                        wsActive.Range("A2:B100").ClearContents
                    End If
                    If fn = False Then
                        MsgBox "No File Selected"
                    Else
                        Sheets("Master").Range("G2").Value = fn
                        Call dural(fn)
                    End If
        With Application
            .DisplayAlerts = True
            .ScreenUpdating = True
        End With
wsActive.Activate
ActiveSheet.Protect "password"

End Sub
Private Sub dural(fn)
   Dim b1 As Workbook, b2 As Workbook
   Dim sh As Worksheet
    Workbooks.Open Filename:=fn
        Set b1 = ActiveWorkbook
        Set b2 = ThisWorkbook
   For Each sh In b1.Sheets
      sh.Copy After:=b2.Sheets(b2.Sheets.Count)
   Next sh
    b1.Close SaveChanges:=False
End Sub
