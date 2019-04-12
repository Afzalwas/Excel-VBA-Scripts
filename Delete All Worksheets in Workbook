Sub Delete_Worksheets()
Dim wsSheet As Worksheet
Set wbBook = ActiveWorkbook
Set wsActive = wbBook.Sheets("Master")
ActiveSheet.Unprotect "password"

'Stopping Application Alerts
Application.DisplayAlerts = False
wsActive.Range("A2:B100").ClearContents

        For Each wsSheet In wbBook.Worksheets
         If wsSheet.Name <> wsActive.Name Then 'And wsSheet.Name <> "Budget_Scenarios_Budgets"
                wsSheet.Activate
                wsSheet.Delete
            End If
        Next wsSheet

'Enabling Application alerts once we are done with our task
ActiveSheet.Protect "password"
Application.DisplayAlerts = True
End Sub
