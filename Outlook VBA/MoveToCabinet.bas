Attribute VB_Name = "MoveToCabinet"
Option Explicit

Sub MoveEmailsToCabinet()

'Instance variables
Dim todayCriteria As String
Dim i As Integer
Dim n As Integer
Dim moveEmail As Boolean
Dim emailsMoved As Integer

'Reference variables
Dim attachment As attachment
Dim toCabinetEmails As items
Dim Outlook As New Outlook.Application
Dim namespace As Outlook.namespace

Set namespace = Application.GetNamespace("MAPI")

'Define the collection filter to restrict processing to emails received before today.
todayCriteria = "[ReceivedTime] < '" & Format(Date, "ddddd h:nn AMPM") & "'"

'Loop through each mailbox in Outlook
    For n = namespace.folders.Count To 1 Step -1
    
    'Determine if the folder is the user's default inbox folder. If so, skip it.
    If namespace.folders(n).folders("Inbox").FolderPath <> namespace.GetDefaultFolder(olFolderInbox).FolderPath Then
        
        'Create the collection of emails with only the old emails
        Set toCabinetEmails = namespace.folders(n).folders("Inbox").items.Restrict(todayCriteria)
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
            toCabinetEmails(i).Move namespace.folders(n).folders("Cabinet")
            emailsMoved = emailsMoved + 1
        End If
        
        'Reset the moveEmail variable to the default of True
        moveEmail = True
        
        Next
    
    Else 'If this is the default mailbox
    
    'for each
    
    
    End If
    
    
Next

MsgBox "All set." & vbNewLine & vbNewLine & emailsMoved & " old emails in your accounts have been moved to their respective cabinets."

End Sub
