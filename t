Sub UpdateEmailsAndExportToExcel()
    Dim objNamespace As Outlook.Namespace
    Dim objRootFolder As Outlook.Folder
    Dim targetEmail As String
    Dim olApp As Outlook.Application
    Dim olExplorer As Outlook.Explorer
    Dim olFolder As Outlook.Folder
    Dim olMail As Outlook.MailItem
    Dim olExcelApp As Object
    Dim olExcelWorkbook As Object
    Dim olExcelSheet As Object
    Dim startTime As Date
    Dim endTime As Date
    
    ' Define the target email address to act on
    targetEmail = "example@example.com"
    
    ' Define the time frame (36 hours ago from now)
    endTime = Now
    startTime = DateAdd("h", -36, endTime) ' Subtract 36 hours
    
    ' Initialize Outlook objects
    Set olApp = Outlook.Application
    Set objNamespace = olApp.GetNamespace("MAPI")
    Set objRootFolder = objNamespace.Folders(targetEmail)
    
    ' Create Excel objects
    Set olExcelApp = CreateObject("Excel.Application")
    Set olExcelWorkbook = olExcelApp.Workbooks.Add
    Set olExcelSheet = olExcelWorkbook.Sheets(1)
    
    ' Create Excel header row
    olExcelSheet.Cells(1, 1).Value = "Subject"
    olExcelSheet.Cells(1, 2).Value = "Received Time"
    
    ' Initialize Outlook Explorer
    Set olExplorer = olApp.ActiveExplorer
    
    ' Loop through all folders inside the Inbox folder
    For Each olFolder In objRootFolder.Folders
        If olFolder.DefaultItemType = olMailItem Then
            ' Loop through the emails in the folder
            For Each olMail In olFolder.Items
                If olMail.ReceivedTime >= startTime And olMail.ReceivedTime <= endTime And olMail.ReplyTime = #1/1/4501# Then
                    ' Email is within the specified time frame and has not been replied to
                    ' Add email details to Excel
                    olExcelSheet.Cells(olExcelSheet.UsedRange.Rows.Count + 1, 1).Value = olMail.Subject
                    olExcelSheet.Cells(olExcelSheet.UsedRange.Rows.Count, 2).Value = olMail.ReceivedTime
                End If
            Next olMail
        End If
    Next olFolder
    
    ' Save and display the Excel workbook
    olExcelWorkbook.SaveAs "C:\Path\To\Save\ExcelWorkbook.xlsx"
    olExcelApp.Visible = True
    olExcelWorkbook.Close
    olExcelApp.Quit
    
    ' Clean up
    Set olExplorer = Nothing
    Set olFolder = Nothing
    Set olMail = Nothing
    Set olExcelSheet = Nothing
    Set olExcelWorkbook = Nothing
    Set olExcelApp = Nothing
    Set objRootFolder = Nothing
    Set objNamespace = Nothing
    Set olApp = Nothing
End Sub
