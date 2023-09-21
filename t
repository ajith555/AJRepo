Option Explicit

Dim objExcelApplication As Object
Dim objExcelWorkbook As Object
Dim objExcelWorksheet As Object
Dim objInbox As Object
Dim targetEmail As String
Dim daysToCheck As Integer

Sub ExportUnrepliedEmails()
    ' Update the target email address and days to check
    targetEmail = "your@email.com"
    daysToCheck = 7
    
    ' Initialize Excel
    Set objExcelApplication = CreateObject("Excel.Application")
    Set objExcelWorkbook = objExcelApplication.Workbooks.Add
    Set objExcelWorksheet = objExcelWorkbook.Worksheets(1)
    
    ' Set up Excel headers
    With objExcelWorksheet
        .Cells(1, 1).Value = "Subject"
        .Cells(1, 2).Value = "Received"
        .Cells(1, 3).Value = "Sender"
        .Cells(1, 4).Value = "Excerpts"
        .Range("A1:D1").Font.Bold = True
    End With
    
    objExcelApplication.Visible = True
    
    ' Initialize Outlook
    Set objInbox = GetOutlookInbox(targetEmail)
    
    If Not objInbox Is Nothing Then
        ' Process emails and subfolders
        ProcessFolder objInbox, daysToCheck
    Else
        MsgBox "Email account not found!", vbExclamation
    End If
    
    ' Format Excel columns and rows
    With objExcelWorksheet
        .Columns("A:C").AutoFit
        .Columns("D").ColumnWidth = 100
        .Columns("D").WrapText = False
    End With
    
    MsgBox "Complete!", vbExclamation
End Sub

Sub ProcessFolder(ByVal objFolder As Object, ByVal days As Integer)
    Dim objItem As Object
    Dim nLastRow As Long
    
    For Each objItem In objFolder.Items
        If objItem.Class = OlObjectClass.olMail Then ' or olMail
            If Not HasReplied(objItem) Then
                If DateDiff("d", objItem.ReceivedTime, Now) <= days Then
                    nLastRow = objExcelWorksheet.Cells(objExcelWorksheet.Rows.Count, 1).End(-4162).Row + 1
                    With objExcelWorksheet
                        .Cells(nLastRow, 1).Value = objItem.Subject
                        .Cells(nLastRow, 2).Value = objItem.ReceivedTime
                        .Cells(nLastRow, 3).Value = objItem.SenderName
                        .Cells(nLastRow, 4).Value = Left(Trim(objItem.Body), 100) & "..."
                    End With
                End If
            End If
        End If
    Next objItem
    
    ' Recursively process subfolders
    If objFolder.Folders.Count > 0 Then
        Dim objSubfolder As Object
        For Each objSubfolder In objFolder.Folders
            ProcessFolder objSubfolder, days
        Next objSubfolder
    End If
End Sub

Function GetOutlookInbox(ByVal email As String) As Object
    On Error Resume Next
    Dim objNamespace As Object
    Dim objRootFolder As Object
    Dim objFolder As Object
    
    Set objNamespace = CreateObject("Outlook.Application").GetNamespace("MAPI")
    Set objRootFolder = objNamespace.Folders(email)
    Set objFolder = objRootFolder.Folders("Inbox")
    
    If Err.Number = 0 Then
        Set GetOutlookInbox = objFolder
    Else
        Set GetOutlookInbox = Nothing
    End If
    On Error GoTo 0
End Function

Function HasReplied(ByVal objMail As Object) As Boolean
    On Error Resume Next
    HasReplied = (objMail.ReplyTime <> 0)
    On Error GoTo 0
End Function
