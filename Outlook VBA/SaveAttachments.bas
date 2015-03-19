Attribute VB_Name = "SaveAttachments"
Option Explicit

Sub SaveAttachments()

    Dim items, Item, myAttachments, myAttachment As Object
    Dim myolApp As New Outlook.Application
    Dim myOlExp As Outlook.Explorer
    Dim selection As Outlook.selection
    Dim Msg As Outlook.MailItem
    Dim FileName As String
    Dim SelectedFolder As Outlook.Folder
    Dim i As Integer
    Dim StoreList As String
    Dim AttachmentCounter As Integer
    Dim otherObject As Excel.Application
    Dim fdFolder As Office.FileDialog
    Dim DirectoryPath As String
    Dim clipboard As MSForms.DataObject
    Set clipboard = New MSForms.DataObject
    Set otherObject = New Excel.Application
    
    Dim SourceNum As Integer
    Dim DestNum As Integer, FileCount As Long
    Dim Temp As String
    Dim AggFile As String
    Dim DateSubmitted As String
    Dim DateasDate As Date
    Dim PPEnd As Date
    Dim FileCounter As Integer
    Dim NumberSaved As Integer

'    otherObject.Visible = False

    Set myOlExp = myolApp.ActiveExplorer
    Set selection = myOlExp.selection

    FileCounter = 1

    With otherObject.Application.FileDialog(msoFileDialogFolderPicker)
        .InitialFileName = "C:\Users\danseethaler\Desktop\"
        .Title = "Select the folder to save the attachments."
        .Show
        If .SelectedItems.Count <> 1 Then Exit Sub
        DirectoryPath = .SelectedItems(1) & "\"
    End With

    otherObject.Quit
    Set otherObject = Nothing
     
    For Each Item In selection
     Set Msg = Item
     
        Set myAttachments = Item.Attachments
        
        If myAttachments.Count > 0 Then
         
        For i = 1 To myAttachments.Count

            If InStr(UCase(myAttachments(i).DisplayName), ".PDF") > 0 Then
                FileName = Left(myAttachments(i).DisplayName, Len(myAttachments(i).DisplayName) - 4) & " " & Replace(FormatDateTime(Msg.SentOn, vbShortDate), "/", "-") & ".pdf"
                
                If Dir(DirectoryPath & FileName) <> "" Then
                    Do Until Dir(DirectoryPath & Left(FileName, Len(FileName) - 4) & "(" & FileCounter & ").pdf") = ""
                    FileCounter = FileCounter + 1
                    Loop
                    FileName = Left(FileName, Len(FileName) - 4) & "(" & FileCounter & ").pdf"
                    FileCounter = 1
                End If
                
                myAttachments(i).SaveAsFile DirectoryPath & FileName
                NumberSaved = NumberSaved + 1
                
            End If
            
        Next i
        
        End If
        
    Next
        
    Set items = Nothing
    Set Item = Nothing
    Set myAttachments = Nothing
    Set myAttachment = Nothing
    Set myolApp = Nothing
    Set myOlExp = Nothing
    Set selection = Nothing

    MsgBox (NumberSaved & " PDF attachments have been saved to " & DirectoryPath)
     
End Sub
