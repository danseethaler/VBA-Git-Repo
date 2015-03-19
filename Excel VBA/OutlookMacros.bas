Attribute VB_Name = "OutlookMacros"
Option Explicit

Sub CopyAndPasteToMailBody()
'https://social.msdn.microsoft.com/Forums/office/en-US/5b4cbfc2-f5f3-4b95-bdc8-25f3eac39965/paste-excel-chart-into-new-outlook-email
'http://www.rondebruin.nl/win/s1/outlook/bmail2.htm
    Set mailApp = CreateObject("Outlook.Application")
    Set mail = mailApp.CreateItem(olMailItem)
    mail.Display
    Set wEditor = mailApp.ActiveInspector.WordEditor
    Selection.Copy
    wEditor.Application.Selection.Paste
    
End Sub


Sub ListOutlookItems(control As IRibbonControl)
Dim AppOutlook As Outlook.Application
Dim strSheet As String
Dim strPath As String
Dim ActiveRow As Integer
Dim ActiveColumn As Integer
Dim Message As Outlook.MailItem
Dim nms As Outlook.Namespace
Dim Folder As Outlook.MAPIFolder
Dim itm As Object
Dim MinDate As Date

MinDate = InputBox("What is the earliest date for which you want emails listed?")

If WorksheetFunction.CountA(Cells) = 0 Then

Set nms = Outlook.Application.GetNamespace("MAPI")
Set Folder = nms.PickFolder

If Folder Is Nothing Then
    MsgBox "There are no mail messages to export", vbOKOnly, _
    "Error"
    Exit Sub
ElseIf Folder.DefaultItemType <> olMailItem Then
    MsgBox "There are no mail messages to export", vbOKOnly, _
    "Error"
    Exit Sub
ElseIf Folder.Items.Count = 0 Then
    MsgBox "There are no mail messages to export", vbOKOnly, _
    "Error"
    Exit Sub
End If

Rows(1).Font.Bold = True

Range("A1").Value = "To"
Range("B1").Value = "From"
Range("C1").Value = "Sender's Email"
Range("D1").Value = "Subject"
Range("E1").Value = "Sent On"

ActiveColumn = 1
ActiveRow = 2

For Each itm In Folder.Items

Set Message = itm

On Error Resume Next

If Message.SentOn > MinDate Then

Cells(ActiveRow, ActiveColumn).Value = Message.To
Cells(ActiveRow, ActiveColumn + 1).Value = Message.Sender
Cells(ActiveRow, ActiveColumn + 2).Value = Message.SenderEmailAddress
If Message.Sender.GetExchangeUser().PrimarySmtpAddressCells <> "" Then Cells(ActiveRow, ActiveColumn + 2).Value = Message.Sender.GetExchangeUser().PrimarySmtpAddress
Cells(ActiveRow, ActiveColumn + 3).Value = Message.Subject
Cells(ActiveRow, ActiveColumn + 4).Value = Message.SentOn

ActiveRow = ActiveRow + 1

End If

Next itm

Else
MsgBox ("This worksheet is not empty. Select an empty worksheet.")

Exit Sub

End If


Call ColumnsAutofitCall

End Sub

Sub CreateFromTemplate(control As IRibbonControl) '
    Dim Cell As Range
    Dim myOlApp As Outlook.Application
    Dim MyItem As Outlook.MailItem
    Dim ItemsSent As Integer
    Dim EmailAction As String
    Dim EmailTemplate As String
    Dim Continue As String
    Dim Header As Range
    Dim Preview As String
    Dim Outlook As Object
    
    Preview = "Yes"

    'Make sure Outlook is open.
    On Error Resume Next
    Set Outlook = GetObject(, "Outlook.Application")
    On Error GoTo 0

    If Outlook Is Nothing Then
        MsgBox "Please open Microsoft Outlook before running this program."
        Exit Sub
    End If

    
    'Check to ensure each cell is filled in the current region.
    If WorksheetFunction.CountA(Range("A1").CurrentRegion) <> Range("A1").CurrentRegion.Count Then
    EmailAction = MsgBox("There are missing values in this list." & vbNewLine & vbNewLine & _
            "Do you want to continue?", vbYesNo)
        If EmailAction = vbNo Then Exit Sub
    End If
    
    'Verify the values in column A are email addresses.
    If IsEmpty(Range("A2")) Then
        MsgBox ("Please enter valid email addresses into column A.")
    Exit Sub
    End If
    
    For Each Cell In Range("A2:A" & Range("A1").End(xlDown).Row)
        If Cell.Value Like "?*@?*.?*" Then
        Else
        MsgBox ("Please enter valid email addresses into " & Cell.Address)
        Exit Sub
        End If
    Next Cell
    
    'Confirm the start of the program and determine if emails will be sent or saved.
    EmailAction = MsgBox("Do you want to send the emails?" & vbNewLine & vbNewLine & _
            "No will save the emails as drafts. Cancel will cancel the program." & vbNewLine & vbNewLine & _
            "Please make sure Outlook is open on your machine.", vbYesNoCancel, "Continue?")
    
    'If the cancel button is selected the program is canceled.
    If EmailAction = vbCancel Then Exit Sub
        
    With Application.FileDialog(msoFileDialogFilePicker)
        If Dir("C:\Users\danseethaler\Dropbox\Work\Incidents\Payroll Processing Error -- Action Required.oft") = "" Then
        .InitialFileName = "C:\Users\danseethaler\Dropbox\Shared at Work\Incidents\Payroll Processing Error -- Action Required.oft"
        Else: .InitialFileName = "C:\Users\danseethaler\Dropbox\Work\Incidents\Payroll Processing Error -- Action Required.oft"
        End If
        .InitialFileName = CreateObject("WScript.Shell").SpecialFolders("Desktop")
        .Title = "Please select an Outlook template."
        .Show
        If .SelectedItems.Count <> 1 Then Exit Sub
        EmailTemplate = .SelectedItems(1)
    End With
    
    'This line prevents the screen from updating while the program is running.
    Application.ScreenUpdating = False
    
    'This line begins the loop from row two through the last row with a value in column A.
    For Each Cell In Range("A2:A" & Range("A1").End(xlDown).Row)
    
    'This line creates the Outlook mail object and assigns it to the designated template.
    Set myOlApp = CreateObject("Outlook.Application")
    'ACTION: Add File Picker to this section
    Set MyItem = myOlApp.CreateItemFromTemplate(EmailTemplate)
    
    'This section manipulates several of the properties of the template to insert
    'the information on the row the program is processing.
    
    MyItem.To = Cell
    'MyItem.BCC = ""
    
    For Each Header In Range("B1:" & Range("A1").End(xlToRight).Address)
        MyItem.HTMLBody = Replace(MyItem.HTMLBody, Header, Cell.Offset(0, Header.Column - 1))
    Next Header
    
    If Preview <> "No" Then
    
    MsgBox ("Please click OK and then review the email message in Outlook to confirm that the message looks correct.")
    
    MyItem.Display
    
    Continue = MsgBox("After reviewing the first email message do you want to continue with the rest of the list?", vbYesNo)
    
    If Continue = vbNo Then
    
    MyItem.Delete
    MsgBox ("This process has been aborted. No emails have been saved or sent.")
    Exit Sub
    
    End If
    MyItem.Close (olSave)
    Preview = "No"
    End If
    
    'This If statement tells the program whether to send or save the email.
    If EmailAction = vbYes Then
        MyItem.Send
    Else
        MyItem.Save
    End If
    
    'The ItemsSent variable simply counts the number of emails generated.
    ItemsSent = ItemsSent + 1
    
    Next Cell
    
    'Screen updating is turned back on.
    Application.ScreenUpdating = True
    
    'This statement returns a message box with the outcome of the program.
    If EmailAction = vbYes Then
        MsgBox (ItemsSent & " emails have been sent.")
    Else
        MsgBox (ItemsSent & " emails have been saved to your drafts folder.")
    End If
    
End Sub

Sub EmailTimeKeepersTemplate(control As IRibbonControl)

    Dim Cell As Range
    Dim myOlApp As Outlook.Application
    Dim MyItem As Outlook.MailItem
    Dim ItemsSent As Integer
    Dim EmailAction As String
    Dim EmailTemplate As String
    Dim Continue As String
    Dim Header As Range
    Dim Preview As String
    Dim Outlook As Object

Dim Friday As String, Tuesday As String, Subject As String, body As String
Dim PP As String
Dim FriDeadline As Date, TuesDeadline As Date

'If Environ("username") <> "danseethaler" Then
'    MsgBox "You do not have access to this tool.", vbCritical, "Access Denied"
'    Exit Sub
'End If

PP = InputBox("Please provide the PP for which time entry is due." & vbNewLine & _
        "This should be a two digit number.")
        
If PP = "" Then Exit Sub

'Set dates
FriDeadline = Date
Do Until Weekday(FriDeadline) = 6
    FriDeadline = FriDeadline + 1
Loop

TuesDeadline = Date + 1
Do Until Weekday(TuesDeadline) = 3
    TuesDeadline = TuesDeadline + 1
Loop

'This line creates the Outlook mail object and assigns it to the designated template.
    Set myOlApp = CreateObject("Outlook.Application")
'ACTION: Add File Picker to this section
    Set MyItem = myOlApp.CreateItemFromTemplate("\\CHQPVUN0066\FINUSR\SHARED\FIN_PYRL\2_Payroll Time & Labor Absence Management\Desk Manual (Information)\Create Time Process Email Template.oft")
    
'This section manipulates several of the properties of the template to insert
'the information on the row the program is processing.
    
    With MyItem
        .Subject = "Create Time Process is Complete for PP" & PP
        .HTMLBody = Replace(MyItem.HTMLBody, "#Friday", Format(FriDeadline, "[$-409]mmmm d, yyyy;@"))
        .HTMLBody = Replace(MyItem.HTMLBody, "#Tuesday", Format(TuesDeadline, "[$-409]mmmm d, yyyy;@"))
        .HTMLBody = Replace(MyItem.HTMLBody, "#PP", "PP" & PP)
        .SentOnBehalfOfName = "GSC-HR-Services@ldschurch.org"
        .Display
    End With

Set myOlApp = Nothing
Set Outlook = Nothing

End Sub

Sub TimeAmericaErrorReport(control As IRibbonControl)
    Dim Cell As Range
    Dim myOlApp As Outlook.Application
    Dim MyItem As Outlook.MailItem
    Dim ItemsSent As Integer
    Dim EmailAction As String
    Dim EmailTemplate As String
    Dim Continue As String
    Dim Header As Range
    Dim Preview As String
    Dim Outlook As Object
    Dim AttachWorkbook As String
    Dim Stores As String
    Dim MissingStores As String
    Dim PP As String
    
    'Make sure Outlook is open.
    On Error Resume Next
    Set Outlook = GetObject(, "Outlook.Application")
    On Error GoTo 0

    If Outlook Is Nothing Then
        MsgBox "Please open Microsoft Outlook before running this program."
        Exit Sub
    End If

    If Dir("\\CHQPVUN0066\FINUSR\SHARED\FIN_PYRL\2_Payroll Time & Labor Absence Management\Desk Manual (Information)\TA100 Uploads Template.oft") <> "" Then
            EmailTemplate = "\\CHQPVUN0066\FINUSR\SHARED\FIN_PYRL\2_Payroll Time & Labor Absence Management\Desk Manual (Information)\TA100 Uploads Template.oft"
        ElseIf Dir("C:\Users\danseethaler\Dropbox\Work\Macros\DI Time Comparison\TA100 Uploads Template.oft") <> "" Then
            EmailTemplate = "C:\Users\danseethaler\Dropbox\Work\Macros\DI Time Comparison\TA100 Uploads Template.oft"
        Else: EmailTemplate = "C:\Users\danseethaler\Dropbox\Shared at Work\Macros\DI Time Comparison\TA100 Uploads Template.oft"
    End If
    
    If IsEmpty(Range("A2")) Then Exit Sub
    
    PP = InputBox("Please enter the two digit pay period.", "Pay Period")
    
    Range("A:A").Value = Range("A:A").Value
    
    ActiveWorkbook.Save
    
    Application.ActiveSheet.Copy
    ActiveWorkbook.SaveAs (CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\PP" & PP & " Load Errors")

    For Each Cell In Range("A2:A" & Range("A1").End(xlDown).Row)

        If Not IsEmpty(Cell) And InStr(1, Stores, Cell) = 0 Then
            Stores = Stores & Cell.Value & "<br>"
        End If
    
    Next Cell
    
    'This line creates the Outlook mail object and assigns it to the designated template.
    Set myOlApp = CreateObject("Outlook.Application")
    'ACTION: Add File Picker to this section
    Set MyItem = myOlApp.CreateItemFromTemplate(EmailTemplate)
    
    'This section manipulates several of the properties of the template to insert
    'the information on the row the program is processing.
    
    With MyItem
        .Attachments.Add ActiveWorkbook.FullName
        .To = "DL-WEL-DIStaff"
        .CC = "danseethaler@ldschurch.org;awilkins@ldschurch.org;DL-GSC-PrcSvc-PR-EmployeeData@ldschurch.org"
        .BCC = ""
        .Subject = .Subject & PP
        .HTMLBody = Replace(MyItem.HTMLBody, "#Stores", Stores)
        .Display
    End With
    
    Application.ActiveWorkbook.Close
    
    Kill (CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\PP" & PP & " Load Errors.xlsx")

MsgBox ("Please run the employee time by TRC report and make corrections to errors.")

End Sub

Sub EmailMissingDIStores(control As IRibbonControl)
    Dim ToList As String
    Dim DirectoryPath As String
    Dim FileName As String
    Dim Cell As Range
    Dim i As Integer
    Dim myOlApp As Outlook.Application
    Dim MyItem As Outlook.MailItem
    
    Dim clipboard As MSForms.DataObject
    Set clipboard = New MSForms.DataObject
    
    Dim Stores As Dictionary
    Set Stores = New Dictionary
    Dim strKey As Variant
    
    DirectoryPath = "\\L12239\CXFUSR\Appl\HR800\PS\Temp\GSC\"

'Instantiate the stores dictionary with email addresses
With Stores
    .Add Key:="American Fork", Item:="OliverMB@ldschurch.org;CarterNP@ldschurch.org"
    .Add Key:="Blackfoot", Item:="Patricia.Fowler@ldschurch.org;David.Dexter@ldschurch.org"
    .Add Key:="Boise", Item:="KayaLL@ldschurch.org;MeredithCH@ldschurch.org"
    .Add Key:="Brigham City", Item:="jrobinette@ldschurch.org;JensenJC@ldschurch.org"
    .Add Key:="Burley", Item:="George.Pethtel@ldschurch.org;SimonsonTK@ldschurch.org"
    .Add Key:="Calimesa", Item:="kumikoeastwood@ldschurch.org;jovany.escobar@ldschurch.org"
    .Add Key:="Cedar City", Item:="kimberlee.jensen@ldschurch.org;David.Stephenson@ldschurch.org"
    .Add Key:="Centerville", Item:="amanda.bawden@ldschurch.org;MoonWW@ldschurch.org"
    .Add Key:="Chula Vista", Item:="mpozo@ldschurch.org;CressallN@ldschurch.org"
    .Add Key:="Colton", Item:="pcampbell@ldschurch.org;MasseyDJ@ldschurch.org"
    .Add Key:="Downtown SLC", Item:="LoseeWe@ldschurch.org;SorensenJE@ldschurch.org"
    .Add Key:="Federal Way", Item:="Barbara.Hellickson@ldschurch.org;HatfieldJJ@ldschurch.org"
    .Add Key:="Harrisville", Item:="MurrayNN@ldschurch.org;ryan.pike@ldschurch.org"
    .Add Key:="Idaho Falls", Item:="jennifer.jensen@ldschurch.org;KelleyAP@ldschurch.org"
    .Add Key:="Las Vegas North", Item:="trshurtleff@ldschurch.org;BondocBL@ldschurch.org"
    .Add Key:="Las Vegas South", Item:="eaguilar@ldschurch.org;mnuttall@ldschurch.org"
    .Add Key:="Layton", Item:="pondmb@ldschurch.org;mechamdw@ldschurch.org"
    .Add Key:="Logan", Item:="FloresMD@familysearch.org;HillRJ@ldschurch.org"
    .Add Key:="Los Angeles", Item:="Sharon.Lamb@ldschurch.org;MeyerDB@ldschurch.org"
    .Add Key:="Mesa", Item:="Carol.Andersen@ldschurch.org;HolmJD@ldschurch.org"
    .Add Key:="Murray", Item:="ulloajime@ldschurch.org;LaudieRD@ldschurch.org"
    .Add Key:="Nampa", Item:="erin.buckley@ldschurch.org;Aaron.J.Pincock@ldschurch.org"
    .Add Key:="Phoenix", Item:="sherri.duke@ldschurch.org;MelzerWL@ldschurch.org"
    .Add Key:="Pocatello", Item:="phay@ldschurch.org;FrancisRD@ldschurch.org"
    .Add Key:="Portland", Item:="harrisonsn@ldschurch.org;GotfredsonSL@ldschurch.org"
    .Add Key:="Preston", Item:="MeidellNB@ldschurch.org;LarsenDG@ldschurch.org"
    .Add Key:="Price", Item:="colleen.byrge@ldschurch.org;brett.stoddart@ldschurch.org"
    .Add Key:="Provo", Item:="NelsonP@ldschurch.org;OlsonWa@ldschurch.org"
    .Add Key:="Rexburg", Item:="tracy.smith@ldschurch.org;GlissmeyerKG@ldschurch.org"
    .Add Key:="Richfield", Item:="alicia.murray@ldschurch.org;BaroneMa@ldschurch.org"
    .Add Key:="Sacramento", Item:="btourtillott@ldschurch.org;ThomasTD@ldschurch.org"
    .Add Key:="Sandy", Item:="stokesmh@ldschurch.org;MontalboMA@ldschurch.org"
    .Add Key:="Seattle", Item:="rscook@ldschurch.org;WestBH@ldschurch.org"
    .Add Key:="St George", Item:="RafterySh@ldschurch.org;BaldwinSD@ldschurch.org"
    .Add Key:="Sugarhouse", Item:="PutnamTJ@ldschurch.org;MaradiagaB@ldschurch.org"
    .Add Key:="Tooele", Item:="sherrywelch@ldschurch.org;jtellez@ldschurch.org;BroadheadCA@ldschurch.org"
    .Add Key:="Tucson", Item:="julie.burke@ldschurch.org;Sherri.Wilson@ldschurch.org"
    .Add Key:="Twin Falls", Item:="DebraMarshall@ldschurch.org;TongeKW@ldschurch.org"
    .Add Key:="Vernal", Item:="jeanne.ruckman@ldschurch.org;KitchenGR@ldschurch.org"
    .Add Key:="Welfare Square", Item:="keslerd@ldschurch.org;MeachamSL@ldschurch.org"
    .Add Key:="West Jordan", Item:="TaylorLL@ldschurch.org;KimmelRJ@ldschurch.org"
    .Add Key:="West Valley", Item:="Phyllis.Doane@ldschurch.org;BagleyBR@ldschurch.org"
End With

    FileName = Dir(DirectoryPath)
    Do While FileName <> ""
        If Stores.Exists(Left(FileName, InStrRev(FileName, ".") - 1)) Then
            Stores.Remove Left(FileName, InStrRev(FileName, ".") - 1)
        Else
            MsgBox "Filename " & FileName & " does not match a member of the stores dictionary."
        End If
        
        FileName = Dir
    Loop

    For Each strKey In Stores.Keys()
        ToList = ToList & Stores(strKey) & ";"
    Next
    
    If Stores.Count = 0 Then
        MsgBox "All DI files have been received."
        Exit Sub
    End If

    Set myOlApp = CreateObject("Outlook.Application")
    Set MyItem = myOlApp.CreateItemFromTemplate("\\chqpvun0066\finusr\SHARED\FIN_PYRL\2_Payroll Time & Labor Absence Management\Desk Manual (Information)\DI TA100 Missing.oft")
    
    With MyItem
        .To = ToList
        .CC = "danseethaler@ldschurch.org;awilkins@ldschurch.org;"
        .BCC = ""
        .Subject = "Missing the PP" & RecentPP() & " Time America TXT File"
        .HTMLBody = Replace(.HTMLBody, "#PP", RecentPP())
        .Display
    End With

End Sub

Sub SendSheet(control As IRibbonControl) '
Dim myOlApp As Outlook.Application
Dim MyItem As Outlook.MailItem
Dim SheetName As String
Dim FileCounter As Integer

Application.ScreenUpdating = False

SheetName = Application.ActiveSheet.Name

    If UCase(Left(SheetName, 5)) = "SHEET" Or UCase(Left(SheetName, 11)) = "NEW_UNSAVED" Then
        SheetName = InputBox("Please provide a name for this attachment.")
    End If

If Dir(CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\" & SheetName & ".xlsx") <> "" Then

        FileCounter = 1
        
        Do Until Dir(CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\" & SheetName & " (" & FileCounter & ").xlsx") = ""
        FileCounter = FileCounter + 1
        Loop

        SheetName = SheetName & " (" & FileCounter & ")"
        
        Application.ActiveSheet.Copy
        ActiveWorkbook.SaveAs FileName:=CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\" & SheetName, FileFormat:=51

    Else
    
        Application.ActiveSheet.Copy
        ActiveWorkbook.SaveAs (CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\" & SheetName)

End If



    'This line creates the Outlook mail object and assigns it to the designated template.
    Set myOlApp = CreateObject("Outlook.Application")
    'ACTION: Add File Picker to this section
    Set MyItem = myOlApp.CreateItem(olMailItem)

    With MyItem
        .Attachments.Add ActiveWorkbook.FullName
        .Subject = SheetName
        .Display
    End With
        
    Application.ActiveWorkbook.Close
    
Kill (CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\" & SheetName & ".xlsx")

Application.ScreenUpdating = True

End Sub

Sub AttachToOpenEmail(control As IRibbonControl)

Dim Continue As String
Dim SheetName As String
Dim outApp As Outlook.Application
Dim OutMail As Outlook.MailItem
Dim FileCounter As Integer

On Error Resume Next
Set outApp = GetObject(, "Outlook.Application")
On Error GoTo 0

If outApp Is Nothing Then
  MsgBox ("Please open MS Outlook before running this macro.")
  Exit Sub
End If

  If outApp.ActiveInspector Is Nothing Then
    
    MsgBox "There is no active email in Outlook. Make sure your draft is expanded in it's own window."
    Exit Sub
    
  End If

Application.ScreenUpdating = False

Continue = MsgBox("Send entire workbook?", vbYesNoCancel)

If Continue = vbCancel Then Exit Sub

If Continue = vbYes Then

    If ActiveWorkbook.Path = vbNullString Then
        If UCase(Left(Application.ActiveSheet.Name, 5)) <> "SHEET" And UCase(Left(Application.ActiveSheet.Name, 5)) <> "NEW_U" Then
            SheetName = Application.ActiveSheet.Name
        Else
            SheetName = InputBox("Please provide a name for this attachment.")
        End If
        
        'Check for existing document and save
        If Dir(CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\" & SheetName & ".xlsx") <> "" Then
    
            Do Until Dir(CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\" & SheetName & "(" & FileCounter & ").xlsx") = ""
            FileCounter = FileCounter + 1
            Loop
    
            SheetName = SheetName & " (" & FileCounter & ")"
    
            ActiveWorkbook.SaveAs (CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\" & SheetName)
    
        Else
    
            ActiveWorkbook.SaveAs (CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\" & SheetName)
    
        End If
    
    Else
    
    ActiveWorkbook.Save
    
    End If
    
End If

If Continue = vbNo Then

    If UCase(Left(Application.ActiveSheet.Name, 5)) <> "SHEET" And UCase(Left(Application.ActiveSheet.Name, 5)) <> "NEW_U" Then

        SheetName = Application.ActiveSheet.Name

    Else

        SheetName = InputBox("Please provide a name for this attachment.")

    End If

        'ActiveWorkbook.Save

    Application.ActiveSheet.Copy

    If Dir(CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\" & SheetName & ".xlsx") <> "" Then

        Do Until Dir(CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\" & SheetName & "(" & FileCounter & ").xlsx") = ""
        FileCounter = FileCounter + 1
        Loop

        SheetName = SheetName & " (" & FileCounter & ")"

        ActiveWorkbook.SaveAs (CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\" & SheetName)

    Else

        ActiveWorkbook.SaveAs (CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\" & SheetName)

    End If
        
End If

With outApp
  If .ActiveInspector Is Nothing Then
    MsgBox "There is no open item"
    Exit Sub
  End If
  
  If Not TypeOf .ActiveInspector.CurrentItem Is MailItem Then
    MsgBox "Type of current item isn't email"
    Exit Sub
  End If
  
  Set OutMail = .ActiveInspector.CurrentItem
  
  If OutMail.Sent Then
    MsgBox "Current email was already sent."
    Exit Sub
  End If
  
  OutMail.Attachments.Add ActiveWorkbook.FullName
  .ActiveInspector.Display
  
End With

Set outApp = Nothing

If Continue = vbNo Then

    ActiveWorkbook.Close False

End If

On Error Resume Next
    Kill CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\" & SheetName & ".xlsx"

Application.ScreenUpdating = True

End Sub
