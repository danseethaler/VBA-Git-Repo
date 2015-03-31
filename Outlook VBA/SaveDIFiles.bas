Attribute VB_Name = "SaveDIFiles"
Option Explicit
    
Sub SaveDITextFileAttachments()
    Dim Continue As String
    Dim holdsAttachment As String
    Dim emailAttachments As Attachments
    Dim DirectoryPath As String
    Dim DirectoryPathDetails As String
    Dim Outlook As New Outlook.Application
    Dim attachmentEmails As items
    Dim FileName As String
    Dim processFolder As Outlook.Folder
    Dim i As Integer
    Dim e As Integer
    Dim StoreList As String
    Dim AttachmentCounter As Integer
    Dim fdFolder As Office.FileDialog
    Dim myDate As Date
    Dim oFS As Object
    Dim strFilename As String
    Dim SourceNum As Integer
    Dim DestNum As Integer, FileCount As Long
    Dim Temp As String
    Dim AggFile As String
    Dim DateSubmitted As Date
    Dim DateSubmittedString As String
    Dim AvgSize As Long
    
    Dim Excel As Excel.Application
    Dim DIWorkbook As Workbook
    Dim WBookName As String
    
    Dim myNamespace As Outlook.Namespace
    Set myNamespace = Application.GetNamespace("MAPI")
    
    'Set process folder to the GSC-DIPayroll Inbox
    On Error Resume Next
        Set processFolder = myNamespace.Folders("GSC-DIPayroll@ldschurch.org").Folders("Inbox")
        Set processFolder = myNamespace.Folders("GSC-DIPayroll").Folders("Inbox")
    On Error GoTo 0
    
    'If the GSC-DIPayroll box doesn't exit then exit the sub
    If processFolder = "" Then
        MsgBox "Please add the GSC-DIPayroll box to your Outlook before running this macro."
        Set myNamespace = Nothing
    End If
    
    'Initiate the collection filters to only process emails with attachments in the GSC-DIPayroll Inbox
    holdsAttachment = "[Attachment] = True"
    Set attachmentEmails = processFolder.items.Restrict(holdsAttachment)
    
    'Sort the attachmentEmails collection by the time they were received
    attachmentEmails.Sort "[ReceivedTime]", True
    
    If (Date - RecentPPDate()) > 4 Then
        Continue = MsgBox("It appears that PP" & RecentPP() & " data has already been processed." & vbNewLine & vbNewLine & _
        "Do you want to continue?", vbYesNo, "Continue?")
        If Continue = vbNo Then Exit Sub
    End If
    
'Check if Excel is open
    On Error Resume Next
    Set Excel = GetObject(, "Excel.Application")
    On Error GoTo 0
    
    'Open MS Excel if it is not already open by creating the Excel object
    If Excel Is Nothing Then Set Excel = CreateObject("Excel.Application")

    If Dir("\\L12239\CXFUSR\Appl\HR800\") = "" Then
    
        With Excel.Application.FileDialog(msoFileDialogFolderPicker)
            .InitialFileName = CreateObject("WScript.Shell").SpecialFolders("Desktop")
            .Title = "Select the folder to save the attachments."
            .Show
            If .SelectedItems.Count <> 1 Then Exit Sub
            DirectoryPath = .SelectedItems(1) & "\"
        End With
    Else
        DirectoryPath = "\\L12239\CXFUSR\Appl\HR800\PS\Temp\GSC\"
    End If

    'Set workbook name and directory path
    WBookName = "DI Email Details - PP" & RecentPP() & ".xlsx"
    
    If Dir("\\CHQPVUN0066\FINUSR\SHARED\FIN_PYRL\") <> "" Then
        DirectoryPathDetails = "\\CHQPVUN0066\FINUSR\SHARED\FIN_PYRL\2_Payroll Time & Labor " & _
            "Absence Management\" & WBookName
        Else
        DirectoryPathDetails = CreateObject("WScript.Shell").SpecialFolders("Desktop") \ " & WBookName"
    End If
    
    
    'If details file already exists
    If Dir(DirectoryPathDetails) <> "" Then
        'Check is workbook is open
        On Error Resume Next
        Set DIWorkbook = Excel.Workbooks(WBookName)
        On Error GoTo 0
            If DIWorkbook Is Nothing Then
            Set DIWorkbook = Excel.Workbooks.Open(DirectoryPathDetails)
            Else
            Excel.Workbooks(WBookName).Activate
            End If
    Else 'If file doesn't exist - create it
        Set DIWorkbook = Excel.Workbooks.Add
        DIWorkbook.SaveAs FileName:=DirectoryPathDetails, FileFormat:=51
    End If

    Excel.ScreenUpdating = False

'If new workbook then set header values and create sheets
If DIWorkbook.Sheets(1).Range("A1").Value = "" Then
    DIWorkbook.Sheets(1).Range("A1").Value = "Store"
    DIWorkbook.Sheets(1).Range("B1").Value = "Sender"
    DIWorkbook.Sheets(1).Range("C1").Value = "Email Address"
    DIWorkbook.Sheets(1).Range("D1").Value = "Subject"
    DIWorkbook.Sheets(1).Range("E1").Value = "Date/Time Sent"
    DIWorkbook.Sheets(1).Range("F1").Value = "File Last Modified"
    DIWorkbook.Sheets(1).Range("G1").Value = "Attachment Size"
    DIWorkbook.Sheets(1).Range("H1").Value = "File Size Variance"
    DIWorkbook.Sheets(1).Range("I1").Value = "Duplicate File?"
    DIWorkbook.Sheets(1).Range("J1").Value = "Wrong Dates?"
    DIWorkbook.Sheets(1).Range("K1").Value = "Invalid Filename?"
    
    DIWorkbook.Sheets(1).Name = "Email Details"
    DIWorkbook.Sheets.Add After:=Excel.Sheets(1)
    DIWorkbook.Sheets(2).Name = "Errors"
    DIWorkbook.Sheets.Add After:=Excel.Sheets(2)
    DIWorkbook.Sheets(3).Name = "File Details"
    DIWorkbook.Sheets(1).Activate
    
End If

    Dim DuplicateFile As Boolean
    Dim InvalidCode As Boolean
    
    For e = attachmentEmails.Count To 1 Step -1
        Set emailAttachments = attachmentEmails(e).Attachments
         
            For i = 1 To emailAttachments.Count
            If InStr(UCase(emailAttachments(i).DisplayName), ".TXT") > 0 Then
         'This case statement associates the first three characters of the file we're processing with
         'the store that uses that three character code. Sometimes stores change the three letter
         'code they use so there may be multiple codes for a single store.
            Select Case Left(UCase(emailAttachments(i).DisplayName), 3)
                Case "AME": FileName = "American Fork.txt": AvgSize = 107608
                Case "BLA": FileName = "Blackfoot.txt": AvgSize = 15915
                Case "BOI": FileName = "Boise.txt": AvgSize = 67874
                Case "BRD": FileName = "Brigham City.txt": AvgSize = 46143
                Case "BRI": FileName = "Brigham City.txt": AvgSize = 46143
                Case "BCD": FileName = "Brigham City.txt": AvgSize = 46143
                Case "BUR": FileName = "Burley.txt": AvgSize = 41448
                Case "CAL": FileName = "Calimesa.txt": AvgSize = 21181
                Case "CED": FileName = "Cedar City.txt": AvgSize = 54222
                Case "CEN": FileName = "Centerville.txt": AvgSize = 99799
                Case "CHU": FileName = "Chula Vista.txt": AvgSize = 51595
                Case "COL": FileName = "Colton.txt": AvgSize = 41093
                Case "TIM": FileName = "DI Manufacturing.txt": AvgSize = 86412
                Case "DWT": FileName = "Downtown SLC.txt": AvgSize = 52737
                Case "FED": FileName = "Federal Way.txt": AvgSize = 38979
                Case "FWE": FileName = "Federal Way.txt": AvgSize = 38979
                Case "FWD": FileName = "Federal Way.txt": AvgSize = 38979
                Case "FW ": FileName = "Federal Way.txt": AvgSize = 38979
                Case "HAR": FileName = "Harrisville.txt": AvgSize = 125467
                Case "IDA": FileName = "Idaho Falls.txt": AvgSize = 71727
                Case "LAS": FileName = "Las Vegas North.txt": AvgSize = 81597
                Case "LVS": FileName = "Las Vegas South.txt": AvgSize = 50297
                Case "LAD": FileName = "Layton.txt": AvgSize = 134096
                Case "LAY": FileName = "Layton.txt": AvgSize = 134096
                Case "LOG": FileName = "Logan.txt": AvgSize = 133043
                Case "LOS": FileName = "Los Angeles.txt": AvgSize = 70357
                Case "MES": FileName = "Mesa.txt": AvgSize = 69928
                Case "MUR": FileName = "Murray.txt": AvgSize = 134351
                Case "NAM": FileName = "Nampa.txt": AvgSize = 26019
                Case "PHO": FileName = "Phoenix.txt": AvgSize = 25577
                Case "POC": FileName = "Pocatello.txt": AvgSize = 48409
                Case "POR": FileName = "Portland.txt": AvgSize = 35220
                Case "PRE": FileName = "Preston.txt": AvgSize = 10354
                Case "PRI": FileName = "Price.txt": AvgSize = 18821
                Case "PRO": FileName = "Provo.txt": AvgSize = 165113
                Case "REX": FileName = "Rexburg.txt": AvgSize = 53118
                Case "RIC": FileName = "Richfield.txt": AvgSize = 28261
                Case "SAC": FileName = "Sacramento.txt": AvgSize = 85400
                Case "SAN": FileName = "Sandy.txt": AvgSize = 110766
                Case "SEA": FileName = "Seattle.txt": AvgSize = 40164
                Case "STG": FileName = "St George.txt": AvgSize = 101296
                Case "SUG": FileName = "Sugarhouse.txt": AvgSize = 109566
                Case "TOO": FileName = "Tooele.txt": AvgSize = 46843
                Case "TUC": FileName = "Tucson.txt": AvgSize = 22202
                Case "TWI": FileName = "Twin Falls.txt": AvgSize = 70358
                Case "VER": FileName = "Vernal.txt": AvgSize = 16416
                Case "VEE": FileName = "Vernal.txt": AvgSize = 16416
                Case "WED": FileName = "Welfare Square.txt": AvgSize = 75728
                Case "WJR": FileName = "West Jordan.txt": AvgSize = 112998
                Case "WVL": FileName = "West Valley.txt": AvgSize = 77570
                Case "WVD": FileName = "West Valley.txt": AvgSize = 77570
                Case Else: FileName = Left(emailAttachments(i).DisplayName, Len(emailAttachments(i).DisplayName) - 4) & ".txt": AvgSize = 1: InvalidCode = True
            End Select
    
            If Dir(DirectoryPath & FileName) <> "" Then DuplicateFile = True
            emailAttachments(i).SaveAsFile DirectoryPath & FileName
                strFilename = DirectoryPath & FileName
                'This creates an instance of the MS Scripting Runtime FileSystemObject class
            Set oFS = CreateObject("Scripting.FileSystemObject")
            AttachmentCounter = AttachmentCounter + 1
            
    ' Open the source text file.
    SourceNum = FreeFile()
    
    Open DirectoryPath & FileName For Input As SourceNum
    
    'Pull the first line of the input file to validate the correct date is being submitted.
        Line Input #SourceNum, Temp
        
        'Extract the date out of the input line and verify it's within the date range.
        DateSubmittedString = InstrLike(Temp, "########")
        DateSubmittedString = Left(Right(DateSubmittedString, 4), 2) & "/" & Right(DateSubmittedString, 2) & "/" & Left(DateSubmittedString, 4)
        DateSubmitted = DateSubmittedString
        
        'If the date reported on the first line of the text file is within the prior pay period dates then clear the date reported
        If DateSubmitted >= RecentPPDate() - 14 And DateSubmitted <= RecentPPDate() Then: DateSubmitted = 0
      
    Close #SourceNum

Dim NextRow As Integer
NextRow = DIWorkbook.Sheets(1).UsedRange.SpecialCells(xlLastCell).Row + 1

    DIWorkbook.Sheets(1).Range("A" & NextRow).Value = Replace(FileName, ".txt", "")
    DIWorkbook.Sheets(1).Range("B" & NextRow).Value = attachmentEmails(e).Sender
    DIWorkbook.Sheets(1).Range("C" & NextRow).Value = attachmentEmails(e).Sender.GetExchangeUser().PrimarySmtpAddress
    DIWorkbook.Sheets(1).Range("D" & NextRow).Value = attachmentEmails(e).Subject
    DIWorkbook.Sheets(1).Range("E" & NextRow).Value = attachmentEmails(e).SentOn
    DIWorkbook.Sheets(1).Range("F" & NextRow).Value = oFS.GetFile(strFilename).Datelastmodified
    DIWorkbook.Sheets(1).Range("G" & NextRow).Value = emailAttachments(i).Size
    DIWorkbook.Sheets(1).Range("H" & NextRow).Value = (emailAttachments(i).Size / AvgSize) - 1
    
    'Highlight the cell color if the variance is greater or less than 30%
    If DIWorkbook.Sheets(1).Range("H" & NextRow).Value > 0.3 Or DIWorkbook.Sheets(1).Range("H" & NextRow).Value < -0.3 Then
        DIWorkbook.Sheets(1).Range("H" & NextRow).Font.Color = -16777024
    End If
    If DuplicateFile = "True" Then
        With DIWorkbook.Sheets(1).Range("I" & NextRow)
            .Value = "Duplicate file."
            .Font.Color = -16777024
        End With
        DuplicateFile = False
    End If
    If DuplicateFile = "True" Then
        With DIWorkbook.Sheets(1).Range("I" & NextRow)
            .Value = "Duplicate file."
            .Font.Color = -16777024
        End With
        DuplicateFile = False
    End If
    If DateSubmitted <> 0 Then
        With DIWorkbook.Sheets(1).Range("J" & NextRow)
            .Value = "Time submitted for " & DateSubmitted
            .Font.Color = -16777024
        End With
        DateSubmitted = 0
    End If
    If InvalidCode = "True" Then
        With DIWorkbook.Sheets(1).Range("K" & NextRow)
            .Value = "Store sent invalid three letter code on file."
            .Font.Color = -16777024
        End With
        InvalidCode = False
    End If
    
    DIWorkbook.Sheets(1).Cells.EntireColumn.AutoFit
    DIWorkbook.Sheets(1).Rows(1).Font.Bold = True
    DIWorkbook.Sheets(1).Columns(7).NumberFormat = "_(* #,##0_);_(* (#,##0);_(* ""-""??_);_(@_)"
    DIWorkbook.Sheets(1).Columns(8).Style = "Percent"
    
    'Move the email to the DIPayroll Cabinet
    attachmentEmails(e).UnRead = False
    
    If processFolder.FolderPath = "\\GSC-DIPayroll@ldschurch.org\Inbox" Then
        attachmentEmails(e).Move myNamespace.Folders("GSC-DIPayroll@ldschurch.org").Folders("Cabinet")
    Else
        attachmentEmails(e).Move myNamespace.Folders("GSC-DIPayroll").Folders("Cabinet")
    End If

            End If
            Next i
    Next
    
    
    'Instatiate the AddIns in Excel (this does not happen automatically when Excel
    'is instatiated programmatically
    Dim CurrAddin As Excel.AddIn
    For Each CurrAddin In Excel.AddIns
    Debug.Print CurrAddin.Name
        If CurrAddin.Installed Then
            CurrAddin.Installed = False
            CurrAddin.Installed = True
        End If
    Next CurrAddin
    
    'Make Excel Visible and turn screen updating back on
    Excel.Visible = True
    Excel.ScreenUpdating = True
    DIWorkbook.Save

    Set emailAttachments = Nothing
    Set Outlook = Nothing
    Set oFS = Nothing
    Set DIWorkbook = Nothing
    Set Excel = Nothing
    
    'Open the GSC folder to display the files that have been loaded
    DirectoryPath = "\\L12239\CXFUSR\Appl\HR800\PS\Temp\GSC\"
    Shell "C:\WINDOWS\explorer.exe """ & DirectoryPath & "", vbNormalFocus

    MsgBox (AttachmentCounter & " attachment(s) have been saved to " & DirectoryPath & "." & vbNewLine & vbNewLine & _
            "The upload details have been saved to the " & DirectoryPathDetails & " file.")
     
     AttachmentCounter = 0
     
End Sub


Sub MoveDIEmails()

    Dim Outlook As New Outlook.Application
    Dim Namespace As Outlook.Namespace
    Dim DestFolder As Outlook.Folder
    Dim cabinetFolder As Outlook.Folder
    Dim attachmentEmails As items
    Dim Item As MailItem
    Dim EmailCount As Integer
    Dim i As Integer
    Dim a As Integer
    Dim afterPP As String
    Dim holdsAttachment As String
    Dim DirectoryPath As String
    Dim Continue As String
    
If (Date - RecentPPDate()) > 4 Then
    Continue = MsgBox("It appears that PP" & RecentPP() & " data has already been processed." & vbNewLine & vbNewLine & _
    "Do you want to continue?", vbYesNo, "Continue?")
    If Continue = vbNo Then Exit Sub
End If
 
 Set Namespace = Application.GetNamespace("MAPI")

'Set the destination folder to the GSC-DIPayroll Inbox
On Error Resume Next
 Set DestFolder = Namespace.Folders("GSC-DIPayroll@ldschurch.org").Folders("Inbox")
 Set DestFolder = Namespace.Folders("GSC-DIPayroll").Folders("Inbox")
On Error GoTo 0

'If the GSC-DIPayroll box doesn't exit then exit the sub
If DestFolder = "" Then
    MsgBox "Please add the GSC-DIPayroll box to your Outlook before running this macro."
    Set Namespace = Nothing
End If

'Set the cabinet folder
On Error Resume Next
 Set cabinetFolder = Namespace.Folders("GSC-DIPayroll@ldschurch.org").Folders("Cabinet")
 Set cabinetFolder = Namespace.Folders("GSC-DIPayroll").Folders("Cabinet")
On Error GoTo 0

    'Initiate the collection filters to only process emails with
    'attachements that were received since the beginning of the last pay period end date
    afterPP = "[ReceivedTime] > '" & Format(RecentPPDate(), "ddddd h:nn AMPM") & "'"
    holdsAttachment = "[Attachment] = True"

Set attachmentEmails = cabinetFolder.items.Restrict(afterPP)
Set attachmentEmails = attachmentEmails.Restrict(holdsAttachment)

For a = attachmentEmails.Count To 1 Step -1
    For i = 1 To attachmentEmails(a).Attachments.Count
        If InStr(UCase(attachmentEmails(a).Attachments(i).DisplayName), ".TXT") > 0 Then
            attachmentEmails(a).Move DestFolder
            EmailCount = EmailCount + 1
            Exit For
        End If
    Next i
Next a
    
    'Validate that all files have been removed from the GSC folder.
    If Dir("\\L12239\CXFUSR\Appl\HR800\PS\Temp\GSC\") <> "" Then
        MsgBox "There appears to be files/folders within the HR800\PS\Temp\GSC folder. This should be emptied " & _
            "out before the files for the current pay period are saved."
        'Open the GSC folder so the extra files can be moved out
        DirectoryPath = "\\L12239\CXFUSR\Appl\HR800\PS\Temp\GSC\"
        Shell "C:\WINDOWS\explorer.exe """ & DirectoryPath & "", vbNormalFocus
    End If

    MsgBox (EmailCount & " email(s) have been moved to the " & DestFolder & " folder.")
    
End Sub

Public Function InstrLike(Strings As String, Pattern As String) As String
Dim k As Long
Dim Answer As String

For k = 1 To Len(Strings) - Len(Pattern)
    If Mid(Strings, k, Len(Pattern)) Like Pattern Then
        InstrLike = Mid(Strings, k, Len(Pattern))
        Exit For
    End If
Next k

End Function


Function RecentPP() As Integer
Dim PP01 As Date
Dim lastPPEndDate As Date

lastPPEndDate = Date - (Date - CDate("8/15/2014")) Mod 14

'Determine the PPEnd date for PP01 of the current year
PP01 = CDate("1/3/2014")
Do Until Year(PP01 + 7) = Year(Date)
    PP01 = PP01 + 14
Loop

If PP01 >= Date Then
    RecentPP = 26
    Exit Function
End If

'Determine which pay period just ended.
RecentPP = 1
Do Until lastPPEndDate = PP01
    PP01 = PP01 + 14
    RecentPP = RecentPP + 1
Loop

End Function

Function RecentPPDate() As Date
RecentPPDate = Date - (Date - CDate("8/15/2014")) Mod 14
End Function
