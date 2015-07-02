VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ItemAddClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents deletedItems As Outlook.items
Attribute deletedItems.VB_VarHelpID = -1
Public WithEvents archivedItems As Outlook.items
Attribute archivedItems.VB_VarHelpID = -1
Public WithEvents casesItems As Outlook.items
Attribute casesItems.VB_VarHelpID = -1
 
Public Sub Initialize_handler()
    
    Set deletedItems = Application.GetNamespace("MAPI").GetDefaultFolder(olFolderDeletedItems).items
    Set archivedItems = Application.GetNamespace("MAPI").folders("danseethaler@ldschurch.org").folders("Archive").items
    Set casesItems = Application.GetNamespace("MAPI").folders("danseethaler@ldschurch.org").folders("Cases").items
 
End Sub


Private Sub deletedItems_ItemAdd(ByVal Item As Object)

Item.UnRead = False
 
End Sub

Private Sub archivedItems_ItemAdd(ByVal Item As Object)

Item.UnRead = False
 
End Sub


Private Sub casesItems_ItemAdd(ByVal Item As Object)

Item.UnRead = False
 
End Sub



