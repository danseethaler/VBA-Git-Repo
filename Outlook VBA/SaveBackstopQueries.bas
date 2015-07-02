Attribute VB_Name = "SaveBackstopQueries"
Option Explicit

Public Sub GetAttachments()
Dim Outlook As Outlook.Application
Dim myNamespace As Outlook.namespace
Dim Inbox As MAPIFolder
Dim Item As MailItem
Dim Atmt As attachment
Dim FileName As String
Dim Deleted As Integer
Dim DirectoryPath As String
Dim FileCounter As Integer

Set Outlook = CreateObject("Outlook.Application")
Set Inbox = Outlook.ActiveExplorer.CurrentFolder

Deleted = 1

Do Until Deleted = 0

Deleted = 0

For Each Item In Inbox.items

    If Item.Attachments.Count = 1 Then
    If Item.Sender.GetExchangeUser().PrimarySmtpAddress = "GLOBALHR-PeopleSoft2@ldschurch.org" Or Item.SenderEmailAddress = "GLOBALHR-PeopleSoft@ldschurch.org" Then
    
        For Each Atmt In Item.Attachments
            
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
            
            If Left(Item.Subject, 6) = "Output" Then
                Atmt.SaveAsFile DirectoryPath & FileName
                Else
                Atmt.SaveAsFile DirectoryPath & Item.Subject & ".xls"
            End If
            
            Item.Delete
            
            Deleted = Deleted + 1
            
            End If
        
        Next
    
    End If
    End If
    
Next

Loop

Set Inbox = Nothing

End Sub
