Attribute VB_Name = "SaveBackstopQueries"
Option Explicit

Public Sub GetAttachments()
Dim myolApp As Outlook.Application
Dim myNamespace As Outlook.namespace
Dim Inbox As MAPIFolder
Dim item As MailItem
Dim Atmt As attachment
Dim FileName As String
Dim Deleted As Integer
Dim DirectoryPath As String
Dim FileCounter As Integer

Set myolApp = CreateObject("Outlook.Application")
Set Inbox = myolApp.ActiveExplorer.CurrentFolder

Deleted = 1

Do Until Deleted = 0

Deleted = 0

For Each item In Inbox.items

    If item.Attachments.Count = 1 Then
    If item.Sender.GetExchangeUser().PrimarySmtpAddress = "GLOBALHR-PeopleSoft2@ldschurch.org" Or item.SenderEmailAddress = "GLOBALHR-PeopleSoft@ldschurch.org" Then
    
        For Each Atmt In item.Attachments
            
            If InStr(UCase(Atmt.DisplayName), ".XLS") > 0 Then
            
            DirectoryPath = CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\Backstop Queries\"
            FileName = Atmt.FileName
            
                If Dir(DirectoryPath & FileName) <> "" Then
                        FileCounter = 1
                    Do Until Dir(DirectoryPath & Left(FileName, Len(FileName) - 4) & "(" & FileCounter & ").pdf") = ""
                    FileCounter = FileCounter + 1
                    Loop
                    FileName = Left(FileName, Len(FileName) - 4) & "(" & FileCounter & ").xls"
                End If
            
            If Left(item.Subject, 6) = "Output" Then
                Atmt.SaveAsFile DirectoryPath & FileName
                Else
                Atmt.SaveAsFile DirectoryPath & item.Subject & ".xls"
            End If
            
            item.Delete
            
            Deleted = Deleted + 1
            
            End If
        
        Next
    
    End If
    End If
    
Next

Loop

Set Inbox = Nothing

End Sub
