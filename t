Option Explicit

Dim objExcelApplication As Excel.Application
Dim objExcelWorkbook As Excel.Workbook
Dim objExcelWorksheet As Excel.Worksheet
Dim objinbox As Outlook.Folder

Sub ExportEmailsNotReplied()
    Dim targetEmail As String
    targetEmail = "kushal-ajith.shetty@ubs.com"
    Dim objNamespace As Outlook.Namespace
    Dim objRootFolder As Outlook.Folder
    Dim objABCFolder As Outlook.Folder
    
    ' Create Excel objects and make Excel visible
    Set objExcelApplication = New Excel.Application
    objExcelApplication.Visible = True

    ' Create a new Excel workbook and worksheet
    Set objExcelWorkbook = objExcelApplication.Workbooks.Add
    Set objExcelWorksheet = objExcelWorkbook.Worksheets(1)
    
    ' Set the row height for all rows in the worksheet
    objExcelWorksheet.Rows.RowHeight = 20
    
    ' Add headers to the Excel worksheet
    With objExcelWorksheet
        .Cells(1, 1) = "Subject"
        .Cells(1, 1).Font.Bold = True
        .Cells(1, 2) = "Received"
        .Cells(1, 2).Font.Bold = True
        .Cells(1, 3) = "Sender"
        .Cells(1, 3).Font.Bold = True
        .Cells(1, 4) = "Recipients"
        .Cells(1, 4).Font.Bold = True
        .Cells(1, 5) = "Excerpts"
        .Cells(1, 5).Font.Bold = True
    End With
    
    ' Get the specified email account and its Inbox folder
    Set objNamespace = Outlook.GetNamespace("MAPI")
    Set objRootFolder = objNamespace.Folders(targetEmail)
    Set objinbox = objRootFolder.Folders("Inbox")
    
    ' Check if the "(1) AGED" folder exists or create it if it doesn't
    Dim foundABC As Boolean
    foundABC = False
    Dim objFolder As Outlook.Folder
    For Each objFolder In objinbox.Folders
        If objFolder.Name = "(1) AGED" Then
            foundABC = True
            Set objABCFolder = objFolder
            Exit For
        End If
    Next objFolder
    
    If Not foundABC Then
        Set objABCFolder = objinbox.Folders.Add("(1) AGED", olFolderInbox)
    End If
    
    ' Check and process the folders to export unreplied emails
    ProcessFolders objinbox, objABCFolder

    ' Format the Excel columns and rows
    With objExcelWorksheet
        .Columns("A:D").AutoFit
        .Columns("E").ColumnWidth = 100
        .Columns("E").WrapText = False
    End With

    ' Move replied emails from "(1) AGED" to "(1) AGED Replied" folder
    MoveRepliedEmails objABCFolder
    
    ' Check for and move emails with duplicate subjects within the "(1) AGED" folder
    MoveDuplicateEmails objABCFolder

    ' Remove duplicates from the Excel workbook
    RemoveDuplicatesFromWorkbook objExcelWorkbook

    MsgBox "Complete!", vbExclamation

    ' Clean up
    Set objExcelWorksheet = Nothing
    Set objExcelWorkbook = Nothing
End Sub

Sub ProcessFolders(ByVal objCurrentfolder As Outlook.Folder, ByVal objDestinationFolder As Outlook.Folder)
    Dim i As Long
    Dim objMail As Outlook.MailItem
    Dim strReplied As String
    Dim nDateDiff As Integer
    Dim nReplyDateDiff As Integer
    Dim nLastRow As Integer
    Dim Recipient As Outlook.Recipient
    
    On Error Resume Next
    
    For i = objCurrentfolder.Items.Count To 1 Step -1
        If objCurrentfolder.Items(i).Class = olMail Then
            Set objMail = objCurrentfolder.Items(i)
            ' Retrieve the "Replied" property value
            strReplied = objMail.propertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x10810003")
            
            ' List of excluded sender names
            Dim excludedSenders As String
            excludedSenders = "Sakaria, Pramod" ' Add other sender names as needed
            If InStr(1, excludedSenders, objMail.SenderName, vbTextCompare) = 0 Then
                ' Check if the email is not replied
                If (Not (strReplied = "102")) And (Not (strReplied = "103")) Then
                    ' Calculate date differences
                    nDateDiff = DateDiff("d", objMail.SentOn, Now)
                    nReplyDateDiff = DateDiff("d", objMail.ReceivedTime, Now)
                
                    ' Check if email is from the last 3 days and not replied for more than 2 days
                    If nDateDiff <= 4 And nReplyDateDiff > 3 And nDateDiff >= 0 And nReplyDateDiff >= 0 Then
                        ' Find the last row in the Excel worksheet
                        nLastRow = objExcelWorksheet.Range("A" & objExcelWorksheet.Rows.Count).End(xlUp).Row + 1
                    
                        ' Populate Excel worksheet with email details
                        With objExcelWorksheet
                            .Cells(nLastRow, 1) = objMail.Subject
                            .Cells(nLastRow, 2) = objMail.ReceivedTime
                            .Cells(nLastRow, 3) = objMail.SenderName
                        End With
                    
                        ' Extract recipient names
                        Dim recipients As String
                        recipients = ""
                        For Each Recipient In objMail.Recipients
                            recipients = recipients & Recipient.Name & "; "
                        Next Recipient
                        recipients = Left(recipients, Len(recipients) - 2) ' Remove trailing "; "
                    
                        ' Populate Excel worksheet with recipients' names and excerpts
                        With objExcelWorksheet
                            .Cells(nLastRow, 4) = recipients
                            .Cells(nLastRow, 5) = Left(Trim(objMail.Body), 100) & "..." ' Excerpts
                        End With

                        ' Move the retrieved email to the "(1) AGED" folder
                        If objCurrentfolder.Name <> objDestinationFolder.Name Then
                            objMail.Move objDestinationFolder
                        End If
                    End If
                End If
            End If
        End If
    Next i

    On Error GoTo 0
    
    If objCurrentfolder.Folders.Count > 0 Then
        Dim objSubfolder As Outlook.Folder
        For Each objSubfolder In objCurrentfolder.Folders
            ProcessFolders objSubfolder, objDestinationFolder
        Next objSubfolder
    End If
End Sub

Sub MoveRepliedEmails(ByVal objSourceFolder As Outlook.Folder)
    Dim objRepliedFolder As Outlook.Folder
    Dim objMail As Outlook.MailItem
    Dim i As Long
    
    ' Check if the "(1) AGED Replied" folder exists or create it if it doesn't
    Dim foundReplied As Boolean
    foundReplied = False
    Dim objFolder As Outlook.Folder
    For Each objFolder In objinbox.Folders
        If objFolder.Name = "(1) AGED Replied" Then
            foundReplied = True
            Set objRepliedFolder = objFolder
            Exit For
        End If
    Next objFolder
    
    If Not foundReplied Then
        Set objRepliedFolder = objinbox.Folders.Add("(1) AGED Replied", olFolderInbox)
    End If

    ' Move replied emails from "(1) AGED" to "(1) AGED Replied" folder
    For i = objSourceFolder.Items.Count To 1 Step -1
        If objSourceFolder.Items(i).Class = olMail Then
            Set objMail = objSourceFolder.Items(i)
            Dim strReplied As String
            strReplied = objMail.propertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x10810003")
            If (strReplied = "102" Or strReplied = "103") Then
                objMail.Move objRepliedFolder
            End If
        End If
    Next i
End Sub

Sub RemoveDuplicatesFromWorkbook(ByVal wb As Excel.Workbook)
    Dim lastRow As Long
    Dim ws As Excel.Worksheet
    Dim rng As Excel.Range
    Dim subjectDict As Object ' Dictionary to track duplicate subjects
    
    Set ws = wb.Worksheets(1)
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Set rng = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, 1)) ' Only consider the "Subject" column (Column A)

    ' Create a Dictionary to track duplicate subjects
    Set subjectDict = CreateObject("Scripting.Dictionary")
    
    Dim subjectCell As Range
    Dim rowIndex As Long
    For rowIndex = lastRow To 2 Step -1 ' Start from the last row and loop backward
        Set subjectCell = rng.Cells(rowIndex, 1)
        If Not subjectCell.Value = "" Then
            If Not subjectDict.Exists(subjectCell.Value) Then
                ' Add subject to the dictionary if it doesn't exist
                subjectDict.Add subjectCell.Value, subjectCell.Row
            Else
                ' If the subject is already in the dictionary, delete the entire row
                ws.Rows(subjectCell.Row).Delete
            End If
        End If
    Next rowIndex
End Sub

Sub MoveDuplicateEmails(ByVal objFolder As Outlook.Folder)
    Dim objDuplicateFolder As Outlook.Folder
    Dim objMail As Outlook.MailItem
    Dim subjectDict As Object ' Dictionary to track duplicate subjects
    Dim i As Long
    
    ' Check if the "(1) AGED Duplicate" folder exists or create it if it doesn't
    Dim foundDuplicate As Boolean
    foundDuplicate = False
    Dim objSubfolder As Outlook.Folder
    For Each objSubfolder In objinbox.Folders
        If objSubfolder.Name = "(1) AGED Duplicate" Then
            foundDuplicate = True
            Set objDuplicateFolder = objSubfolder
            Exit For
        End If
    Next objSubfolder
    
    If Not foundDuplicate Then
        Set objDuplicateFolder = objinbox.Folders.Add("(1) AGED Duplicate", olFolderInbox)
    End If

    ' Create a Dictionary to track duplicate subjects
    Set subjectDict = CreateObject("Scripting.Dictionary")
    
    For i = objFolder.Items.Count To 1 Step -1
        If objFolder.Items(i).Class = olMail Then
            Set objMail = objFolder.Items(i)
            Dim strReplied As String
            strReplied = objMail.propertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x10810003")
            If (strReplied <> "102" And strReplied <> "103") Then ' Only consider unreplied emails
                Dim subject As String
                subject = objMail.Subject
                If Not subjectDict.Exists(subject) Then
                    subjectDict.Add subject, objMail.EntryID ' Track the subject and its EntryID
                Else
                    ' If the subject is already in the dictionary, move the email to the "(1) AGED Duplicate" folder
                    objMail.Move objDuplicateFolder
                End If
            End If
        End If
    Next i
End Sub














Title: Simplify Email Management with a Clever VBA Trick

Keeping up with emails can feel like a wild ride, right? But what if there's a smart way to make it easier? Say hello to a neat VBA trick that sorts, organizes, and tidies up your emails—just the way you want!

Introduction

Juggling emails is a bit like managing a busy puzzle. But there's a cool trick that can make things smoother. This VBA trick is like your own email helper, and it's all about saving time and staying on top of important messages.

Getting to Know the Trick

Here's the magic: a VBA trick that buddies up Excel and Outlook. These two work together to figure out which emails are important to you and then do what you want with them—almost like having a mini assistant!

Awesome Benefits

Smart Sorting: Imagine a filter that's really smart. This trick goes through your inbox and follows rules you set. Like, who sent it, whether you replied, and when it arrived. It sorts things out so you don't have to deal with the mess.

Easy Tracking: Have you ever wanted emails to tell you if you replied? This trick does that for you! It remembers which emails need your attention.

Excel Magic: The trick talks to Excel and makes cool summaries of your emails. It's like making a neat list of what's been going on. Super handy for keeping track.

No More Copycats: Ever get the same email more than once? This trick stops that. It spots those emails and takes care of them—no repeats allowed!

Your Rules, Your Way: You're in charge! You can change the rules any time you like. It's all about what works best for you.

How It Works

The trick uses Outlook and Excel together. It checks your emails, follows your rules, and puts emails where you say. Plus, it's like a detective—it finds important stuff and makes cool reports in Excel. That way, you see all the action in one place.

Who Can Use It?

Office All-Stars: Managers and leaders, this one's for you. The trick helps sort emails super fast, so you never miss out on important stuff.

Customer Heroes: If you're in sales or support, this trick can be your sidekick. It makes sure you see customer messages right away.

Project Masters: For those in charge of projects, this trick keeps you organized. No more lost project updates in your inbox.

Wrap It Up

Bringing this clever VBA trick into your email world is like having an email helper by your side—without the fuss! It takes care of the email chaos so you can focus on chatting and getting things done.

Remember: To try out this trick, you might need to know a little about VBA, Outlook, and Excel. And don't forget about privacy and safety when using it.










CAN BE USED By


Business Executives and Managers:
As a business executive or manager, you receive a substantial amount of emails daily. This macro can help you identify emails that require urgent attention, track unanswered emails, and organize communication with your team and clients. It ensures that you never miss an important email and enables you to maintain effective communication.

Sales and Customer Support Teams:
Sales representatives and customer support teams can use the macro to manage client interactions. It assists in prioritizing customer emails, tracking unresolved queries, and ensuring swift responses. This leads to improved customer satisfaction and streamlined communication.

Project Managers:
Project managers often deal with a high volume of emails related to project updates, tasks, and deadlines. The macro can aid in categorizing project-related emails, tracking progress, and identifying critical messages that require immediate attention. This helps maintain project timelines and keeps team members informed.

Team Collaboration:
Within a team, members can use the macro to collaborate effectively. It helps in monitoring team discussions, sharing project updates, and tracking the status of different tasks. The macro ensures that everyone stays informed and engaged.

Event Planners:
Event planners can benefit from the macro by organizing event-related communications. It assists in keeping track of vendor inquiries, client requests, and logistics discussions. The macro's filtering and categorization capabilities simplify the coordination process.

Freelancers and Entrepreneurs:
Individuals who manage their businesses or work as freelancers can leverage the macro to streamline client communication. It aids in managing project proposals, contract discussions, and client follow-ups. The macro ensures that important details are easily accessible and helps maintain professional relationships.

Inbox Cleanup and Organization:
If your inbox is frequently cluttered with emails, this macro can help you declutter and organize it. It identifies and moves emails to specific folders based on your criteria, reducing inbox overload and making it easier to find important messages.

Monitoring Important Alerts:
For individuals who receive important alerts or notifications via email, the macro can automatically prioritize and categorize these messages. This ensures that critical alerts are not buried in the inbox and allows for swift action.

Deadline and Task Management:
The macro can aid in managing deadlines and tasks by categorizing emails related to specific projects or tasks. It helps prevent missing important deadlines and keeps tasks on track.

Personal Email Management:
Even for personal email accounts, the macro can be useful. It can help in identifying and categorizing emails from friends, family, social groups, and newsletters, making it easier to manage personal correspondence.
