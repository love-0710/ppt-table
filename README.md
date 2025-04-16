# PPT Table
This code will help you to make a ppt table that contain, 
<br>
S/N |	Author(s) |	Paper & Publication Details |	Findings |  Research Gap  |	Relevance to the Project






#code

Dim outlookApp, mailItem, attachmentPath, fso

On Error Resume Next

' Create Outlook Application
Set outlookApp = CreateObject("Outlook.Application")
If outlookApp Is Nothing Then
    MsgBox "Outlook is not installed or cannot be accessed.", vbCritical
    WScript.Quit
End If

' Create Mail Item
Set mailItem = outlookApp.CreateItem(0)
If mailItem Is Nothing Then
    MsgBox "Failed to create mail item.", vbCritical
    WScript.Quit
End If

' Set attachment path
attachmentPath = "C:\Users\e720312\Documents\scripts\sample.pdf"

' Check file existence
Set fso = CreateObject("Scripting.FileSystemObject")
If Not fso.FileExists(attachmentPath) Then
    MsgBox "Attachment file does not exist: " & attachmentPath, vbCritical
    WScript.Quit
End If

' Add attachment
mailItem.To = "receiver@example.com"
mailItem.Subject = "Automated Report"
mailItem.Body = "Please find the attached document."

mailItem.Attachments.Add attachmentPath
If Err.Number <> 0 Then
    MsgBox "Error attaching file: " & Err.Description, vbCritical
    Err.Clear
    WScript.Quit
End If

' Show email (change to .Send if needed)
mailItem.Display

On Error GoTo 0














Sub SendMailWithAttachment()
    Dim outlookApp As Object
    Dim mailItem As Object

    Set outlookApp = CreateObject("Outlook.Application")
    Set mailItem = outlookApp.CreateItem(0)

    mailItem.To = "receiver@example.com"
    mailItem.Subject = "Test Subject"
    mailItem.Body = "Please find the attached file."

    mailItem.Attachments.Add "C:\Users\e720312\Documents\scripts\sample.pdf"
    mailItem.Display
End Sub
