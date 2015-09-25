Attribute VB_Name = "SaveQAQueries"
Option Explicit

Public Sub GetAttachments()

'Setting Variables for later reference in code
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

'Setup the dictionary of queries
'Make sure to Click Tools > References... > Then click
'The Microsoft Scripting Runtime Library

Dim dicKey As String
Dim newQueries As New Scripting.Dictionary
Dim dict As New Scripting.Dictionary
Dim keyy As Variant
Dim dictString As String

dict.Add "AW_COMPFREQCHECK", 1
dict.Add "AW_COMPFREQCHECK2", 1
dict.Add "HOURLY_MONTHLYCOMP", 1
dict.Add "AW_TERM_NO_REASONCODE", 1
dict.Add "SWTS_TAX", 1
dict.Add "LOC_TAX_AUDIT_COUNTY", 1
dict.Add "TL_STAT_WORKGROUP", 1
dict.Add "CELL_PHONE_REIMB_QA", 1
dict.Add "AW_GSCPARTTIMEWORKGROUP", 1
dict.Add "AW_STDBU", 1
dict.Add "EMP_NO_SSN", 1
dict.Add "JOB_POSITION_DONT_MATCH", 1
dict.Add "AW_HIRED_MULTIPLE_EMPL_RCD", 1
dict.Add "GSC_REPORTS_TO_BLANK", 1
dict.Add "AW_PSD_NEXEO_QUALITY", 1
dict.Add "SUT_TAX_2", 1
dict.Add "GSC_POSITIONS_CHANGEDHRS", 1
dict.Add "EAF_NOTPROCESSED", 1
dict.Add "AW_CWRPAYGROUP", 1
dict.Add "INTERNATIONAL_REGION_USA", 1
dict.Add "AW_HOURLYTOSALARY", 1
dict.Add "AW_DAILYHIRES", 1
dict.Add "AW_ICSCWRPAYRATES", 1
dict.Add "AW_DOUBLEPAY", 1
dict.Add "DIPAYDISTRIBUTIONREHIRE_QA", 1
dict.Add "AW_BAC_POSITION_CHANGES", 1
dict.Add "AWQA_I9_COMP_MISMATCH", 1
dict.Add "AW_PARTTOFULL", 1
dict.Add "AW_ICSCWRWG", 1
dict.Add "AW_INTERNAT_PAY_SYS", 1
dict.Add "DIPAYDISTRIBUTIONTERMDAILY_QA", 1
dict.Add "AWQA_CES_EMPLCLASS", 1
dict.Add "GSC_DEPTID_TASKPROFILE_COMP", 1
dict.Add "AWQA_I9_SSN_MISMATCH", 1
dict.Add "INCORRECT_BENEFIT_SER_ERRORREP", 1
dict.Add "AW_ALTERNATEBENEFIT", 1

Dim queryEmails As New collection

Dim namespace As Outlook.namespace
Set namespace = Application.GetNamespace("MAPI")

'For each email in the current user's inbox
For Each Item In namespace.GetDefaultFolder(olFolderInbox).items

    'Check to see if the email has just one attachment and that the email was sent by GlobalHR
    If InStr(1, Item.Subject, "QA Query Flag") > 0 Then
        QueryCount = QueryCount + 1
        Item.UnRead = False
        'Add each email from GlobalHR which has one attachment to the queryEmails collection
        queryEmails.Add Item
    End If
    
Next Item

If QueryCount < 1 Then
    Set namespace = Nothing
    MsgBox "There were no emails with the 'QA Query Flag'."
    Exit Sub
End If

QueryCount = 0

'Set the directory path to the employee's desktop
DirectoryPath = CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\"

'AJ - comment out the above line and uncomment the below line
'DirectoryPath = "M:\HRPS Team\Backstop Queries\Archive" & "\"

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
        
        'Turn of the display of errors in case there is no excel workbook with this name to open
        On Error Resume Next
        
        'Check is workbook is open
        Set WBook = Excel.Workbooks(WBookName)
        
        'Turn the display of errors back on
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
    
'Loop through every email (MailItem) in the primary inbox
For Each Item In queryEmails
        For Each Atmt In Item.Attachments
            If InStr(UCase(Atmt.DisplayName), ".XLS") > 0 Then
                FileName = Atmt.FileName
                
                'Set the dicKey variable to part of the filename
                dicKey = Left(FileName, InStr(1, FileName, "-") - 1)
                
                'Remove the query from the dictionary
                If dict.Item(dicKey) Then
                    dict.Remove (dicKey)
                Else
                    newQueries.Add dicKey, 1
                End If
            
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
                
                'Create the variable FullPath by adding the previously defined variables of DirectoryPath (my backstop query folder) and FileName (the name of the together
                FullPath = DirectoryPath & FileName

            End If
            
            'Save the Excel file to my backstop query folder
            Atmt.SaveAsFile FullPath
            
            'Open the Excel file just saved
            Excel.Workbooks.Open FullPath
                'Check to see if the query has any data and if not delete it
                If Excel.Workbooks(FileName).Sheets(1).Range("B1") = " 0" Then
                    Excel.Workbooks(FileName).Close
                    Kill FullPath
                Else
                'If the query has data, name the new tab with the first 30 charaters of the subject since there is a character limit(34) to the name of a tab in Excel
                Excel.Workbooks(FileName).Sheets(1).Name = Left(FileName, 30)
                If Left(Item.Subject, 6) = "Output" Then
                    Excel.Workbooks(FileName).Sheets(1).Range("C1").Value = Item.Body
                    Else
                    Excel.Workbooks(FileName).Sheets(1).Range("C1").Value = Item.Subject
                End If
                'Add the new tab after the previous one in the workbook
                Excel.Workbooks(FileName).Sheets(1).Move After:=Excel.Workbooks(WBookName).Sheets(1)
                Kill FullPath
                
                ValidQueryCount = ValidQueryCount + 1
                    
                End If
            
            QueryCount = QueryCount + 1
            Item.Delete
        Next
Next

If Excel.Workbooks(WBookName).Sheets(1).Name = "Sheet1" Then _
Excel.Workbooks(WBookName).Sheets(1).Delete

With Excel
    .Workbooks(WBookName).Save
    .Visible = True
    .EnableEvents = True
    .ScreenUpdating = True
End With

MsgBox QueryCount & " backstop query emails have been processed. " & ValidQueryCount & " queries actually had data."

'Check for any remaining dictionary keys
If dict.Count > 0 Then

    For Each keyy In dict
        dictString = dictString + keyy + ", "
   Next
    
    MsgBox "Missing the following " + CStr(dict.Count) + " queries: " + dictString
End If

'Check for new queries in the newQueries dictionary
If newQueries.Count > 0 Then

    dictString = ""

    For Each keyy In newQueries
        dictString = dictString + keyy + ", "
    Next
    
    MsgBox "Please add the following " + CStr(newQueries.Count) + " queries to the hard coded list: " + dictString
End If

Set namespace = Nothing
Set WBook = Nothing
Set Excel = Nothing

End Sub
