Option Explicit

Dim objExcelApplication As Excel.Application
Dim objExcelWorkbook As Excel.Workbook
Dim objExcelWorksheet As Excel.Worksheet

Sub ExportEmailsNotReplied()
    Dim targetEmail As String
    targetEmail = "testing@ubs.com"
    
    Set objExcelApplication = CreateObject("Excel.Application")
    Set objExcelWorkbook = objExcelApplication.Workbooks.Add
    Set objExcelWorksheet = objExcelWorkbook.Worksheets(1)
    
    With objExcelWorksheet
        .Cells(1, 1) = "Subject"
        .Cells(1, 1).Font.Bold = True
        .Cells(1, 2) = "Received"
        .Cells(1, 2).Font.Bold = True
        .Cells(1, 3) = "Sender"
        .Cells(1, 3).Font.Bold = True
        .Cells(1, 4) = "Excerpts"
        .Cells(1, 4).Font.Bold = True
    End With
    
    objExcelApplication.Visible = True
    
    ' Get the specified email account and its Inbox folder
    Dim objNamespace As Outlook.Namespace
    Dim objRootFolder As Outlook.Folder
    Set objNamespace = Outlook.GetNamespace("MAPI")
    Set objRootFolder = objNamespace.Folders(targetEmail)
    
    ' Check and process the folders
    ProcessFolders objRootFolder
    
    ' Format the Excel columns and rows
    With objExcelWorksheet
        .Columns("A:C").AutoFit
        .Columns("D").ColumnWidth = 100
        .Columns("D").WrapText = False
    End With
    
    RemoveDuplicatesFromWorkbook
    
    MsgBox "Complete!", vbExclamation
End Sub

Sub ProcessFolders(ByVal objCurrentfolder As Outlook.Folder)
    Dim i As Long
    Dim objMail As Outlook.MailItem
    Dim strReplied As String
    Dim nDateDiff As Integer
    Dim nLastRow As Integer
    
    For i = objCurrentfolder.Items.Count To 1 Step -1
        If objCurrentfolder.Items(i).Class = olMail Then
            Set objMail = objCurrentfolder.Items(i)
            strReplied = objMail.propertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x10810003")
            
            If (Not (strReplied = "102")) And (Not (strReplied = "103")) Then
                nDateDiff = DateDiff("d", objMail.SentOn, Now)
                
                If nDateDiff < 7 Then
                    nLastRow = objExcelWorksheet.Range("A" & objExcelWorksheet.Rows.Count).End(xlUp).Row + 1
                    
                    With objExcelWorksheet
                        .Cells(nLastRow, 1) = objMail.Subject
                        .Cells(nLastRow, 2) = objMail.ReceivedTime
                        .Cells(nLastRow, 3) = objMail.SenderName
                        .Cells(nLastRow, 4) = Left(Trim(objMail.Body), 100) & "..."
                    End With
                End If
            End If
        End If
    Next i
    
    If objCurrentfolder.Folders.Count > 0 Then
        Dim objSubfolder As Outlook.Folder
        For Each objSubfolder In objCurrentfolder.Folders
            ProcessFolders objSubfolder
        Next objSubfolder
    End If
End Sub

Sub RemoveDuplicatesFromWorkbook()
    Dim lastRow As Long
    Dim lastCol As Long
    Dim ws As Worksheet
    Dim rng As Range
    
    Set ws = objExcelWorkbook.Worksheets(1)
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    Set rng = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))
    rng.RemoveDuplicates Columns:=Array(1, 2, 3, 4), Header:=xlYes
End Sub
