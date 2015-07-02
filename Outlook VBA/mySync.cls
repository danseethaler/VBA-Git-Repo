VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "mySync"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim WithEvents mySync As Outlook.syncObject
Attribute mySync.VB_VarHelpID = -1
 
Sub Initialize_handler()
 Set mySync = Application.Session.SyncObjects.Item(1)
 mySync.Start
End Sub
 
Private Sub mySync_SyncEnd()
 'MsgBox Application.GetNamespace("MAPI").GetDefaultFolder(olFolderInbox).items.Restrict("[UnRead] = True").Count & _
    " Emails are unread in the main inbox."
End Sub
