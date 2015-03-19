Attribute VB_Name = "SaveQAQueries"
Option Explicit

Public Sub GetAttachments()

Dim Item As MailItem
Dim Atmt As attachment
Dim FileName As String
Dim QueryCount As Integer
Dim ValidQueryCount As Integer
Dim DirectoryPath As String
Dim FileCounter As Integer
Dim WBookName As String
Dim WBook As Workbook
Dim Excel As Excel.Application
Dim FullPath As String

Dim Namespace As Outlook.Namespace
Set Namespace = Application.GetNamespace("MAPI")

For Each Item In Namespace.GetDefaultFolder(olFolderInbox).items
If Item.Attachments.Count = 1 And Item.SenderName = "GLOBALHR-PeopleSoft" Then QueryCount = QueryCount + 1
Next Item

If QueryCount < 1 Then
Set Namespace = Nothing
Exit Sub
End If

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

    With Excel
        .EnableEvents = False
        .ScreenUpdating = False
    End With
    
'*********************************
'Loop through every email (MailItem) in the primary inbox
For Each Item In Namespace.GetDefaultFolder(olFolderInbox).items
    If Item.Attachments.Count = 1 And Item.SenderName = "GLOBALHR-PeopleSoft" Then
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
                Excel.Workbooks(FileName).Sheets(1).Move After:=Excel.Workbooks(WBookName).Sheets(1)
                Kill FullPath
                
                ValidQueryCount = ValidQueryCount + 1
                    
                End If
            
            Item.Delete
        Next

    End If
Next

With Excel
    .Workbooks(WBookName).Save
    .Visible = True
    .EnableEvents = True
    .ScreenUpdating = True
End With

MsgBox QueryCount & " backstop query emails have been processed. " & ValidQueryCount & " queries actually had data."

Set Namespace = Nothing
Set WBook = Nothing
Set Excel = Nothing

End Sub
