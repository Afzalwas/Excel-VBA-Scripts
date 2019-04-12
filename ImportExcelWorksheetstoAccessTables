Sub ImportAccess()
Dim appAccess As Access.Application

Set wbBook = ActiveWorkbook
Set wsActive = wbBook.Sheets("Master")
  
    ThisWorkbookPath = ActiveWorkbook.Path & "/" & ActiveWorkbook.Name

        Set appAccess = CreateObject("Access.Application")
        appAccess.Visible = True

            NewDatabaseFilePath = ActiveWorkbook.Path & "/" & "New File Name.mdb"

                If FileExists(NewDatabaseFilePath) Then
                    Kill NewDatabaseFilePath
                    appAccess.NewCurrentDatabase _
                        FilePath:=NewDatabaseFilePath, _
                        FileFormat:=acNewDatabaseFormatAccess2002
                Else
                    appAccess.NewCurrentDatabase _
                        FilePath:=NewDatabaseFilePath, _
                        FileFormat:=acNewDatabaseFormatAccess2002
                End If

                        For Each wsSheet In wbBook.Worksheets
                            If wsSheet.Name <> wsActive.Name Then
                                
                                ImportSheetName = wsSheet.Name
                                ImportSheetName2 = ImportSheetName & "!"
                                
                                    appAccess.docmd.TransferSpreadsheet _
                                        TransferType:=acImport, _
                                        SpreadsheetType:=acSpreadsheetTypeExcel12Xml, _
                                        TableName:=ImportSheetName, _
                                        Filename:=ThisWorkbookPath, _
                                        HasFieldNames:=True, _
                                        Range:=ImportSheetName2
                                End If
                        Next wsSheet
                        
                            Dim db As DAO.Database
                            Dim tdf As DAO.TableDef
                            Set db = CurrentDb
                            
                        For Each tdf In db.TableDefs
     
                        Next

End Sub

Function FileExists(ByVal FileToTest As String) As Boolean
   FileExists = (Dir(FileToTest) <> "")
End Function
