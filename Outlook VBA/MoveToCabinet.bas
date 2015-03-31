Attribute VB_Name = "MoveToCabinet"
Option Explicit

Sub MoveToCabinet()

'Instance variables
Dim todayCriteria As String
Dim i As Integer
Dim moveEmail As Boolean
Dim emailsMoved As Long

'Reference variables
Dim attachment As attachment
Dim toCabinetEmails As items
Dim emailAccount As Account
Dim accountInbox As Outlook.Folder
Dim Outlook As New Outlook.Application
Dim Namespace As Outlook.Namespace
Dim destFolder As Outlook.Folder

Set Namespace = Application.GetNamespace("MAPI")

'Define the collection filter to restrict processing to emails received before today.
todayCriteria = "[ReceivedTime] < '" & Format(Date, "ddddd h:nn AMPM") & "'"

'Loop through each mailbox in Outlook
For Each emailAccount In Namespace.Accounts
    
    'Set the accountInbox reference variable. This can be done in one of two ways
    'depending on how the user setup the account
    On Error Resume Next
     Set accountInbox = Namespace.Folders(emailAccount.DisplayName).Folders("Inbox")
     Set accountInbox = Namespace.Folders(emailAccount.UserName).Folders("Inbox").items.Restrict(todayCriteria)
    On Error GoTo 0

    'Set the destFolder reference variable. This can be done in one of two ways
    'depending on how the user setup the account
    On Error Resume Next
     Set destFolder = Namespace.Folders(emailAccount.DisplayName).Folders("Cabinet")
     Set destFolder = Namespace.Folders(emailAccount.UserName).Folders("Cabinet").items.Restrict(todayCriteria)
    On Error GoTo 0
    
    'Determine if the folder is the user's default inbox folder. If so, skip it.
    If accountInbox.FolderPath <> Namespace.GetDefaultFolder(olFolderInbox).FolderPath Then
        
        Set toCabinetEmails = accountInbox.items.Restrict(todayCriteria)
        toCabinetEmails.Sort "[ReceivedTime]", True
            
        For i = toCabinetEmails.Count To 1 Step -1
            toCabinetEmails(i).UnRead = False
            
            If toCabinetEmails(i).Attachments.Count > 0 Then
            
                For Each attachment In toCabinetEmails(i).Attachments
                
                    'Only move the emails that don't have a .wav attachment (these emails don't create cases
                    If LCase(Right(attachment.FileName, 4)) = ".wav" Then
                            moveEmail = False
                            Exit For
                        Else
                            moveEmail = True
                    End If
                    
                Next
                
            Else
            
                'If there are no attachments set moveEmail to true
                moveEmail = True
                
            End If
        
        If moveEmail Then
            toCabinetEmails(i).Move destFolder
            emailsMoved = emailsMoved + 1
        End If
        
        'Reset the moveEmail variable to the default of True
        moveEmail = True
        
        Next
        
    End If
    
    
Next

MsgBox "All set." & vbNewLine & emailsMoved & " old emails in your accounts have been moved to their respective cabinets."

End Sub
