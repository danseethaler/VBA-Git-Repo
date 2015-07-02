Attribute VB_Name = "PersonalTools"
Option Explicit

Sub GenerateEmail()
'For Testing
Dim outApp As Object
Dim OutMail As Object

If Environ("username") <> "danseethaler" Then
    MsgBox "You do not have access to this tool.", vbCritical, "Access Denied"
    Exit Sub
End If

Set outApp = CreateObject("Outlook.Application")
Set OutMail = outApp.CreateItem(0)

With OutMail
    .To = ""
    .SentOnBehalfOfName = InputBox("Who would you like this email to come from?", "Sent on Behalf of Name", "GSC-HR-Services@ldschurch.org")
    .Subject = ""
    .body = ""
    .Display
End With

Set OutMail = Nothing
Set outApp = Nothing

End Sub

Sub EmailTimeKeepers() 'control As IRibbonControl

Dim outApp As Object
Dim OutMail As Object
Dim Friday As String, Tuesday As String, Subject As String, body As String
Dim PP As String
Dim FriDeadline As Date, TuesDeadline As Date

If Environ("username") <> "danseethaler" Then
    MsgBox "You do not have access to this tool.", vbCritical, "Access Denied"
    Exit Sub
End If

PP = InputBox("Please provide the PP for which time entry is due." & vbNewLine & _
        "This should be a two digit number.")
        
If PP = "" Then Exit Sub

Subject = "Create Time Process is Complete for PP" & PP

With Application
    .EnableEvents = True
    .ScreenUpdating = True
End With

'Set dates
FriDeadline = Date
Do Until Weekday(FriDeadline) = 6
    FriDeadline = FriDeadline + 1
Loop

TuesDeadline = Date + 1
Do Until Weekday(TuesDeadline) = 3
    TuesDeadline = TuesDeadline + 1
Loop


Set outApp = CreateObject("Outlook.Application")
Set OutMail = outApp.CreateItem(0)

body = "The 'Create Time' process is now complete for PP" & PP & "." & vbNewLine & vbNewLine & _
        "The standard deadline for time entry is 5:00pm this coming Friday, " & Format(FriDeadline, "[$-409]mmmm d, yyyy;@") & ". " & _
        "Managers should approve reported time and absence events by 5:00pm on Friday, " & Format(FriDeadline, "[$-409]mmmm d, yyyy;@") & ". " & _
        "Time sheet corrections can be made until time entry is closed at 10:00 a.m. on Tuesday, " & Format(TuesDeadline, "[$-409]mmmm d, yyyy;@") & ". " & _
        "Due to the time constraints of processing payroll we will strictly enforce this 10:00am deadline." & vbNewLine & vbNewLine & _
        "We encourage all employees to be prompt in reporting time and ask managers to take time to review and approve reported time. " & _
        "Off-cycle checks are often the result of time not being entered correctly and managers failing to review time for their employees." & vbNewLine & vbNewLine & _
        "Time Keeper Note:  Each pay period a PeopleSoft process is run to automatically generate time for active employees whose time is entered by a time keeper. " & _
        "These are non-self-service employees who do not use external time keeping systems (i.e. Kronos, Time America, etc.). " & _
        "The purpose of the 'Create Time' process is to reduce the amount of redundant time entry required by time keepers. " & _
        "This process loads each time sheet with the standard hours on the employee's schedule. " & _
        "After the process is complete each time keeper should review the standard hours and update the time sheet to reflect actual hours worked." & vbNewLine & vbNewLine & _
        "Self-Service Note:  Self-service employees can enter time anytime (except during payroll processing) and are encouraged to enter time regularly. " & _
        "This helps ensure time is accurate when managers are ready to approve time prior to payroll processing." & vbNewLine & vbNewLine & _
        "Please note:  All times listed are Salt Lake City time." & vbNewLine & vbNewLine & _
        "Thanks for all you do to ensure that employees are paid correctly." & vbNewLine & vbNewLine & _
        "If you are no longer responsible for time entry or should no longer receive these email reminders, " & _
        "please let us know and we will remove you from the time entry email distribution list." & vbNewLine & vbNewLine & _
        "HR Processing Services" & vbNewLine & _
        "GSC - Payroll Services" & vbNewLine & _
        "(801) 240-8080" & vbNewLine & _
        "GSC-HR-Services@ldschurch.org"

With OutMail
    .To = ""
    .SentOnBehalfOfName = "GSC-HR-Services@ldschurch.org"
    .CC = ""
    .BCC = "OaksS@ldschurch.org; awilkins@ldschurch.org; Time Keepers Email DL; FM/Distribution Email DL"
    .Subject = Subject
    .body = body
    .Display
End With
On Error GoTo 0

With Application
    .EnableEvents = True
    .ScreenUpdating = True
End With

Set OutMail = Nothing
Set outApp = Nothing

End Sub

Sub ListAddIns()
Dim a As AddIn
With Application
    For Each a In .AddIns2
    Debug.Print a.Path & "\" & a.Name & " " & a.Installed
    Next
End With
End Sub

Sub QuickSaveDesktop(control As IRibbonControl)
Dim SheetName As String
Dim FileCounter As Integer
Dim Continue As String
Dim OriginalFile As String

Application.ScreenUpdating = False

Continue = MsgBox("Save entire workbook?", vbYesNoCancel)

If Continue = vbCancel Then Exit Sub

If Continue = vbNo Then

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
        ActiveWorkbook.SaveAs fileName:=CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\" & SheetName, FileFormat:=51

    Else
    
        Application.ActiveSheet.Copy
        ActiveWorkbook.SaveAs (CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\" & SheetName), FileFormat:=51

End If

ActiveWorkbook.Close

End If

If Continue = vbYes Then

SheetName = ActiveWorkbook.Name

    If ActiveWorkbook.Path = vbNullString Or ActiveWorkbook.Path = Replace(CreateObject("WScript.Shell").SpecialFolders("Desktop"), "Desktop", "Downloads") Then
    
    SheetName = InputBox("What would you like to call this workbook?")
    
    If SheetName = "" Then Exit Sub
        
        'Check for existing document
        If Dir(CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\" & SheetName) <> "" Then
            Do Until Dir(CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\" & SheetName & "(" & FileCounter & ").xlsx") = ""
                FileCounter = FileCounter + 1
            Loop
            
            SheetName = SheetName & " (" & FileCounter & ")"
        End If

    ActiveWorkbook.SaveAs (CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\" & SheetName), FileFormat:=51
    
    Else
    
    OriginalFile = ActiveWorkbook.FullName
    ActiveWorkbook.Save
    
        'Check for existing document and save
        If Dir(CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\" & SheetName) <> "" Then
        FileCounter = 1
            Do Until Dir(CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\" & Left(SheetName, InStrRev(SheetName, ".") - 1) & " (" & FileCounter & ").xlsx") = ""
                FileCounter = FileCounter + 1
            Loop
            
            SheetName = Left(SheetName, InStrRev(SheetName, ".") - 1) & " (" & FileCounter & ")"
        End If
        
    ActiveWorkbook.SaveAs (CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\" & SheetName), FileFormat:=51
    ActiveWorkbook.Close
    Workbooks.Open fileName:=OriginalFile
    
    End If

End If

Application.ScreenUpdating = True

End Sub

Sub ListGUIDsForReferenceLibraries()
     'Macro purpose:  To determine full path and Globally Unique Identifier (GUID)
     'to each referenced library.  Select the reference in the Tools\References
     'window, then run this code to get the information on the reference's library

Workbooks.Add

    Dim i As Long
    On Error Resume Next
    With ActiveWorkbook.Sheets(1)
        .Range("A1") = "Reference name"
        .Range("B1") = "Full path to reference"
        .Range("C1") = "Reference GUID"
    End With
    For i = 1 To ThisWorkbook.VBProject.References.Count
        With ThisWorkbook.VBProject.References(i)
            ActiveWorkbook.Sheets(1).Range("A65536").End(xlUp).Offset(1, 0) = .Name
            ActiveWorkbook.Sheets(1).Range("A65536").End(xlUp).Offset(0, 1) = .FullPath
            ActiveWorkbook.Sheets(1).Range("A65536").End(xlUp).Offset(0, 2) = .GUID
        End With
    Next i
    On Error GoTo 0

End Sub
