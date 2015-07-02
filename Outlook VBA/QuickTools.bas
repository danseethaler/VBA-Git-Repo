Attribute VB_Name = "QuickTools"
Option Explicit

Sub CountSelectedEmails()
    Dim outlookExplorer As Outlook.Explorer
    Dim selection As Outlook.selection
    Dim Item As MailItem
    
    Set outlookExplorer = Outlook.ActiveExplorer
    Set selection = outlookExplorer.selection
    
    If selection.Count = 1 Then
    
            MsgBox "This email is in the " & selection(1).Parent & " folder."
    
        Else
        
            MsgBox selection.Count & " items are currently selected."
    
    End If

End Sub


Sub NewEmailToRecipients()
    Dim currentEmail As MailItem
    Dim outlookExplorer As Outlook.Explorer
    Dim selection As Outlook.selection
    Dim newEmail As MailItem
    
    Set outlookExplorer = Outlook.ActiveExplorer
    Set selection = outlookExplorer.selection
    
    Set currentEmail = selection.Item(1)
    
    Debug.Print currentEmail.Sender
    
    Set newEmail = CreateItem(olMailItem)
    
    If currentEmail.Sender <> "Dan Seethaler" Then
        newEmail.To = Replace(currentEmail.To, "Dan Seethaler", currentEmail.Sender)
    Else
        newEmail.To = currentEmail.To
    End If
    
    newEmail.CC = currentEmail.CC
    newEmail.Display

End Sub
