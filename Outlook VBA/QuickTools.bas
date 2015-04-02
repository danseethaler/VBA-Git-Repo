Attribute VB_Name = "QuickTools"
Option Explicit

Sub CountSelectedEmails()
    Dim outlookExplorer As Outlook.Explorer
    Dim selection As Outlook.selection
    Dim item As MailItem
    
    Set outlookExplorer = Outlook.ActiveExplorer
    Set selection = outlookExplorer.selection
    
    If selection.Count = 1 Then
    
            MsgBox "This email is in the " & selection(1).Parent & " folder."
    
        Else
        
            MsgBox selection.Count & " items are currently selected."
    
    End If

End Sub

Sub MarkDeletedItemsRead()
Dim myNamespace As Outlook.namespace
Dim deletedEmails As items
Dim item As Object
Dim i As Integer

    'Set the deletedEmails collection to the unread emails in the default Deleted Items folder
    Set myNamespace = Application.GetNamespace("MAPI")
    Set deletedEmails = myNamespace.GetDefaultFolder(olFolderDeletedItems).items.Restrict("[UnRead] = True")
    
    'Iterate through the emails in reverse and change the UnRead property to False
    For i = deletedEmails.Count To 1 Step -1
        deletedEmails(i).UnRead = False
    Next

End Sub
