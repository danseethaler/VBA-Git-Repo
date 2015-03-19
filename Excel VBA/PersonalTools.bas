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

Sub AppendFiles(control As IRibbonControl) '
Dim SourceNum As Integer
Dim DestNum As Integer, FileCount As Long
Dim Temp As String, DirectoryPath As String
Dim FileName As String, Cell As Range
Dim AggFile As String
Dim AddName As String
Dim FMACounter As Integer

AggFile = InputBox("Please provide a name for the aggregated file.", "File Name") & ".txt"

If AggFile = ".txt" Then GoTo EndEarly

'Identify the folder that contains the files to merge
    With Application.FileDialog(msoFileDialogFolderPicker)
        Select Case Dir("\\L12239\CXFUSR\Appl\HR800\PS\Temp\GSC\", vbDirectory)
        Case "": .InitialFileName = "C:\Users\danseethaler\Dropbox\Shared\Macros\DI Time Comparison\TXT Files\"
        Case Else: .InitialFileName = "\\L12239\CXFUSR\Appl\HR800\PS\Temp\GSC\"
        End Select
        .Title = "Select the folder with texts files to merge."
        .Show
        
    Select Case .SelectedItems.Count
        Case Is = 0: GoTo EndEarly
        Case Is = 1: DirectoryPath = .SelectedItems(1) & "\"
    End Select
    
    End With
    
AddName = MsgBox("Do you want to append the file name to each line of data?", vbYesNo)
    
    FileName = Dir(DirectoryPath, vbReadOnly) ' + vbHidden)
    
Application.ScreenUpdating = False
    
Do While FileName <> ""
    Application.StatusBar = FileCount & " Files Complete..."
    If FileName <> AggFile And InStr(UCase(FileName), ".TXT") > 0 Then
      ' Open the destination text file.
      DestNum = FreeFile()
      Open DirectoryPath & AggFile For Append As DestNum

      ' Open the source text file.
      SourceNum = FreeFile()
      Open DirectoryPath & FileName For Input As SourceNum

      ' Include the following line if the first line of the source
      ' file is a header row that you do now want to append to the
      ' destination file:
      ' Line Input #SourceNum, Temp

      ' Read each line of the source file and append it to the
      ' destination file.
      Do While Not EOF(SourceNum)
         Line Input #SourceNum, Temp
         If AddName = vbYes Then
         Print #DestNum, Temp & "," & FileName
         Else: Print #DestNum, Temp
         If InStr(1, Temp, "FMA") Then: FMACounter = FMACounter + 1
         End If
      Loop
      
        FileName = Dir
        FileCount = FileCount + 1
        
      Close #DestNum
      Close #SourceNum
      
    Else: FileName = Dir

    End If
        
    Loop

EndEarly:

Application.ScreenUpdating = True
Application.StatusBar = False

    MsgBox FileCount & " file(s) have been aggregated into the " & Chr(34) & AggFile & Chr(34) & " file."
    If FMACounter > 0 Then: MsgBox ("There are " & FMACounter & " lines that contain the FMA TRC. Please contact " & _
    "the store to resolve this.")
End Sub

Sub ConcatenateDelimitedText(control As IRibbonControl) '
Dim ConcatRange As Range
Dim Cell As Range

Set ConcatRange = Intersect(Selection, Columns(ActiveCell.Column))

For Each Cell In ConcatRange

Select Case Selection.Columns.Count
    Case 11: Cell.Value = Cell & " " & Cell.Offset(0, 1) & " " & Cell.Offset(0, 2) & " " & Cell.Offset(0, 3) & " " & Cell.Offset(0, 4) & " " & Cell.Offset(0, 5) & " " & Cell.Offset(0, 6) & " " & Cell.Offset(0, 7) & " " & Cell.Offset(0, 8) & " " & Cell.Offset(0, 9) & " " & Cell.Offset(0, 10)
    Case 10: Cell.Value = Cell & " " & Cell.Offset(0, 1) & " " & Cell.Offset(0, 2) & " " & Cell.Offset(0, 3) & " " & Cell.Offset(0, 4) & " " & Cell.Offset(0, 5) & " " & Cell.Offset(0, 6) & " " & Cell.Offset(0, 7) & " " & Cell.Offset(0, 8) & " " & Cell.Offset(0, 9)
    Case 9: Cell.Value = Cell & " " & Cell.Offset(0, 1) & " " & Cell.Offset(0, 2) & " " & Cell.Offset(0, 3) & " " & Cell.Offset(0, 4) & " " & Cell.Offset(0, 5) & " " & Cell.Offset(0, 6) & " " & Cell.Offset(0, 7) & " " & Cell.Offset(0, 8)
    Case 8: Cell.Value = Cell & " " & Cell.Offset(0, 1) & " " & Cell.Offset(0, 2) & " " & Cell.Offset(0, 3) & " " & Cell.Offset(0, 4) & " " & Cell.Offset(0, 5) & " " & Cell.Offset(0, 6) & " " & Cell.Offset(0, 7)
    Case 7: Cell.Value = Cell & " " & Cell.Offset(0, 1) & " " & Cell.Offset(0, 2) & " " & Cell.Offset(0, 3) & " " & Cell.Offset(0, 4) & " " & Cell.Offset(0, 5) & " " & Cell.Offset(0, 6)
    Case 6: Cell.Value = Cell & " " & Cell.Offset(0, 1) & " " & Cell.Offset(0, 2) & " " & Cell.Offset(0, 3) & " " & Cell.Offset(0, 4) & " " & Cell.Offset(0, 5)
    Case 5: Cell.Value = Cell & " " & Cell.Offset(0, 1) & " " & Cell.Offset(0, 2) & " " & Cell.Offset(0, 3) & " " & Cell.Offset(0, 4)
    Case 4: Cell.Value = Cell & " " & Cell.Offset(0, 1) & " " & Cell.Offset(0, 2) & " " & Cell.Offset(0, 3)
    Case 3: Cell.Value = Cell & " " & Cell.Offset(0, 1) & " " & Cell.Offset(0, 2)
    Case 2: Cell.Value = Cell & " " & Cell.Offset(0, 1)

    Case Else: MsgBox ("Only 9 Columns can be concatenated with this function."): Exit Sub

End Select

Range(Cell.Offset(0, 1), Cell.Offset(0, Selection.Columns.Count - 1)).Clear

    Do Until Right(Cell, 1) <> " "
        Cell = Left(Cell, Len(Cell) - 1)
    Loop

Next

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
        ActiveWorkbook.SaveAs FileName:=CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\" & SheetName, FileFormat:=51

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
    Workbooks.Open FileName:=OriginalFile
    
    End If

End If

Application.ScreenUpdating = True

End Sub

Sub ImportExternalTimeFile(control As IRibbonControl)
Dim DirectoryPath As String
Dim FileName As String

With Application.FileDialog(msoFileDialogFolderPicker)
        .InitialFileName = CreateObject("WScript.Shell").SpecialFolders("Desktop")
        .Title = "Select the folder with external time files."
        .Show
        
    Select Case .SelectedItems.Count
        Case Is = 0: Exit Sub
        Case Is = 1: DirectoryPath = .SelectedItems(1) & "\"
    End Select
    
End With

Application.ScreenUpdating = False

Sheets.Add

'Set headers
    Range("A1").FormulaR1C1 = "File Name"
    Range("B1").FormulaR1C1 = "EmplID"
    Range("C1").FormulaR1C1 = "EmplRcd"
    Range("D1").FormulaR1C1 = "Report Date"
    Range("E1").FormulaR1C1 = "TRC"
    Range("F1").FormulaR1C1 = "Hours"
    Range("G1").FormulaR1C1 = "Amount"
    Range("H1").FormulaR1C1 = "Profile"
    Range("I1").FormulaR1C1 = "Business Unit"
    Range("J1").FormulaR1C1 = "Deptid"
    Range("K1").FormulaR1C1 = "Account"
    Range("L1").FormulaR1C1 = "Product"
    Range("M1").FormulaR1C1 = "Project ID"
    Range("N1").FormulaR1C1 = "Business Unit PC"

Range("B2").Select

FileName = Dir(DirectoryPath, vbReadOnly) ' + vbHidden)

    Do While FileName <> ""

If UCase(Right(FileName, 4)) = ".TXT" Or UCase(Right(FileName, 4)) = ".DAT" Then
    
    With ActiveSheet.QueryTables.Add(Connection:= _
        "TEXT;" & DirectoryPath & FileName _
        , Destination:=Range(ActiveCell.Address))
        '.CommandType = 0
        .Name = Left(FileName, Len(FileName) - 4)
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .TextFilePromptOnRefresh = False
        .TextFilePlatform = 437
        .TextFileStartRow = 1
        .TextFileParseType = xlFixedWidth
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .TextFileConsecutiveDelimiter = False
        .TextFileTabDelimiter = True
        .TextFileSemicolonDelimiter = False
        .TextFileCommaDelimiter = False
        .TextFileSpaceDelimiter = False
        
            'List of fields from external time files.
            '$emplid:11 $emplrcd:3 $reportdate:10 $trc:5 $hours:6 $amt:8 $profile:1 $business_unit:5 $deptid:10
            '$account:6 $product:6 $project_id:15 $business_unit_pc:5 $activity_id:15 $resource_type:5
            '$resource_cat:5 $resource_sub_cat:5
            
        .TextFileColumnDataTypes = Array(1, 1, 3, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1)
        .TextFileFixedColumnWidths = Array(11, 3, 10, 5, 6, 8, 1, 5, 10, 6, 6, 15, 5, 15, 5, 5, 5)
        .TextFileTrailingMinusNumbers = True
        .Refresh BackgroundQuery:=False
    End With

Range(ActiveCell.Offset(0, -1).Address, ActiveCell.End(xlDown).Offset(0, -1)) = Left(FileName, Len(FileName) - 4)

Range("B" & ActiveSheet.UsedRange.SpecialCells(xlLastCell).Row + 1).Select

End If

FileName = Dir

Loop

    Range("A1").CurrentRegion.Copy
    Sheets.Add
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    Application.DisplayAlerts = False
    Sheets(ActiveSheet.Index + 1).Delete
    Application.DisplayAlerts = True
    
'Call ConvertEmpIDToTextCall(Range("B:B"))

Columns("D:D").NumberFormat = "m/d/yyyy"
ActiveSheet.Name = "External Files PP" & InputBox("What is the two digit pay period?")

Range("O:O").Delete
Range("O:O").Delete
Range("O:O").Delete
Range("O:O").Delete

Columns.AutoFit

    Range("A1").Select

End Sub

Sub ImportDLFiles()
Dim DirectoryPath As String
Dim FileName As String

With Application.FileDialog(msoFileDialogFolderPicker)
        .InitialFileName = CreateObject("WScript.Shell").SpecialFolders("Desktop")
        .Title = "Select the folder with external time files."
        .Show
        
    Select Case .SelectedItems.Count
        Case Is = 0: Exit Sub
        Case Is = 1: DirectoryPath = .SelectedItems(1) & "\"
    End Select
    
End With

FileName = Dir(DirectoryPath, vbReadOnly) ' + vbHidden)

    Do While FileName <> ""

If UCase(Right(FileName, 4)) = ".TXT" Or UCase(Right(FileName, 4)) = ".DAT" Then

    With ActiveSheet.QueryTables.Add(Connection:= _
        "TEXT;" & DirectoryPath & FileName _
        , Destination:=Range(ActiveCell.Address))
        .Name = "TIMEKEEPER A-D"
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .TextFilePromptOnRefresh = False
        .TextFilePlatform = 437
        .TextFileStartRow = 1
        .TextFileParseType = xlDelimited
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .TextFileConsecutiveDelimiter = True
        .TextFileTabDelimiter = True
        .TextFileSemicolonDelimiter = False
        .TextFileCommaDelimiter = False
        .TextFileSpaceDelimiter = True
        .TextFileColumnDataTypes = Array(1, 1, 1, 1)
        .TextFileTrailingMinusNumbers = True
        .Refresh BackgroundQuery:=False
    End With

Range("A" & ActiveSheet.UsedRange.SpecialCells(xlLastCell).Row + 1).Select

End If

FileName = Dir

Loop

Columns.AutoFit

    Range("A1").Select

End Sub

Sub ListReferencePaths()
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
