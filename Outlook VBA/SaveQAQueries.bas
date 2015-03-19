Attribute VB_Name = "SaveQAQueries"
Option Explicit

Public Sub GetAttachments()

Dim Item As MailItem
Dim Atmt As Attachment
Dim FileName As String
Dim FileCounter As Integer
Dim QueryCount As Integer
Dim ValidQueryCount As Integer
Dim DirectoryPath As String
Dim FullPath As String
Dim WBookName As String
Dim WBook As Workbook
Dim Excel As Excel.Application
Dim queryEmails As New Collection

Dim Namespace As Outlook.Namespace
Set Namespace = Application.GetNamespace("MAPI")

For Each Item In Namespace.GetDefaultFolder(olFolderInbox).Items
If Item.Attachments.Count = 1 And Item.SenderName = "GLOBALHR-PeopleSoft" Then
QueryCount = QueryCount + 1
queryEmails.Add Item
End If
Next Item

If QueryCount < 1 Then
Set Namespace = Nothing
Exit Sub
End If

QueryCount = 0

DirectoryPath = CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\"

'-Initialize Excel Objects-

'Set workbook name
    WBookName = "Backstop Queries " & Replace(Date, "/", "-") & ".xlsx"
'Check if Excel is open
    On Error Resume Next
    Set Excel = GetObject(, "Excel.Application")
    On Error GoTo 0
'Set Excel Object
    If Excel Is Nothing Then Set Excel = CreateObject("Excel.Application")

'Check if the file in question already exists
    If Dir(DirectoryPath & WBookName) <> "" Then
        'Check is workbook is open
        On Error Resume Next
        Set WBook = Excel.Workbooks(WBookName)
        On Error GoTo 0
            If WBook Is Nothing Then
            Set WBook = Excel.Workbooks.Open(DirectoryPath & WBookName)
            Else
            Excel.Workbooks(WBookName).Activate
            End If
    Else 'If file doesn't exist - create it
        Set WBook = Excel.Workbooks.Add
        WBook.SaveAs FileName:=DirectoryPath & WBookName, FileFormat:=51
    End If

    Excel.ScreenUpdating = False
    
'Loop through every email (MailItem) in the primary inbox
For Each Item In queryEmails
        For Each Atmt In Item.Attachments
            If InStr(UCase(Atmt.DisplayName), ".XLS") > 0 Then
            FileName = Atmt.FileName
            
                'If the file already exists in the directory path give it a unique name
                If Dir(DirectoryPath & FileName) <> "" Then
                        FileCounter = 1
                    Do Until Dir(DirectoryPath & Left(FileName, Len(FileName) - 4) & "(" & FileCounter & ").pdf") = ""
                    FileCounter = FileCounter + 1
                    Loop
                    FileName = Left(FileName, Len(FileName) - 4) & "(" & FileCounter & ").xls"
                End If

'**********************************
'Aggregate Files with Data
                
                FullPath = DirectoryPath & FileName

            End If
            
            Atmt.SaveAsFile FullPath
            
            Excel.Workbooks.Open FullPath
                If Excel.Workbooks(FileName).Sheets(1).Range("B1") = " 0" Then
                    Excel.Workbooks(FileName).Close
                    Kill FullPath
                Else
                Excel.Workbooks(FileName).Sheets(1).Name = Left(FileName, 30)
                If Left(Item.Subject, 6) = "Output" Then
                    Excel.Workbooks(FileName).Sheets(1).Range("C1").Value = Item.Body
                    Else
                    Excel.Workbooks(FileName).Sheets(1).Range("C1").Value = Item.Subject
                End If
                Excel.Workbooks(FileName).Sheets(1).Move After:=Excel.Workbooks(WBookName).Sheets(1)
                Kill FullPath
                
                ValidQueryCount = ValidQueryCount + 1
                    
                End If
            
            Debug.Print Item.Subject
            QueryCount = QueryCount + 1
            Item.Delete
        Next
Next

If Excel.Workbooks(WBookName).Sheets(1).Name = "Sheet1" Then _
Excel.Workbooks(WBookName).Sheets(1).Delete

MsgBox QueryCount & " backstop query emails have been processed. " & ValidQueryCount & " queries actually had data."

Excel.Workbooks(WBookName).Save
Excel.Visible = True
Excel.ScreenUpdating = True

Set Namespace = Nothing
Set WBook = Nothing
Set Excel = Nothing

End Sub
