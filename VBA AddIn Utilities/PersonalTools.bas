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

Sub AppendFiles(control As IRibbonControl)
'This macro will aggregate all the files with extension .txt in the selected folder
'into a new file. This is used to aggregate the external time files from Time America
'into a single file that can be uploaded a single time rather than uploading dozens
'of seperate files. All original files remain unaltered.

Dim SourceNum As Integer
Dim DestNum As Integer, FileCount As Long
Dim Temp As String, DirectoryPath As String
Dim FileName As String, cell As Range
Dim AggFile As String
Dim AddName As String
Dim FMACounter As Integer

'Determine the filename for the new file.
AggFile = InputBox("Please provide a name for the aggregated file.", "File Name") & ".txt"

'If the user did not type anything in for a filename then the macro will terminate
If AggFile = ".txt" Then Exit Sub

'Ask the user for the folder that contains the files to merge
    With Application.FileDialog(msoFileDialogFolderPicker)
        Select Case Dir("\\L12239\CXFUSR\Appl\HR800\PS\Temp\GSC\", vbDirectory)
        Case "": .InitialFileName = "C:\Users\danseethaler\Dropbox\Shared\Macros\DI Time Comparison\TXT Files\"
        Case Else: .InitialFileName = "\\L12239\CXFUSR\Appl\HR800\PS\Temp\GSC\"
        End Select
        .Title = "Select the folder with texts files to merge."
        .Show
    
    'If no folder is selected then terminate the macro.
    'Otherwise set the DirectoryPath variable equal to the path of the selected folder.
    Select Case .SelectedItems.Count
        Case Is = 0: Exit Sub
        Case Is = 1: DirectoryPath = .SelectedItems(1) & "\"
    End Select
    
    End With

'Ask the user if they would like the filename appended to each line of data.
'This is used to determine which line of data goes with which file.
'You should choose No when preparing the data to load into PeopleSoft.
'Select Yes if the data has been loaded and you're appending the files a second time for comparison purposes.
Dim config As Integer
config = vbYesNoCancel + vbDefaultButton2
AddName = MsgBox("Do you want to append the file name to each line of data?", config)

If AddName = vbCancel Then Exit Sub

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

      ' Read each line of the source file and append it to the destination file.
      Do While Not EOF(SourceNum)
           Line Input #SourceNum, Temp
           
           'If we need to append the filename to the line then do it.
           If AddName = vbYes Then
           Print #DestNum, Temp & "," & Left(FileName, Len(FileName) - 4)
           
           'Otherwise just add the source line to the destination file.
           Else: Print #DestNum, Temp
           End If
           
           'Check if the line has FML reported (a common error in DI Files)
           'We used to contact the store about this before we loaded the file
           'Now we simply delete this error on the PeopleSoft error page in the
           'run control and let the store send an email to update the timesheet if needed.
           If InStr(1, Temp, "FMA") Then: FMACounter = FMACounter + 1
           
        Loop
        
        'Get the next file in the directory an increase the number of files processed by one.
          FileName = Dir
          FileCount = FileCount + 1
        
      'Close the files we've been using. The new file will be opened in the next loop.
      Close #DestNum
      Close #SourceNum
    
    'If the file does not have a .txt extension then move onto the next file.
    Else: FileName = Dir

    End If
        
    Loop

'Reset the status bar and begin updating the screen again.
Application.ScreenUpdating = True
Application.StatusBar = False

'Send a message box to the user indicating what happend in the macro.
    MsgBox FileCount & " file(s) have been aggregated into the " & Chr(34) & AggFile & Chr(34) & " file."
    If FMACounter > 0 Then: MsgBox ("There are " & FMACounter & " lines that contain the FMA TRC. Please contact " & _
    "the store to resolve this.")
    
End Sub

Sub ConcatenateDelimitedText(control As IRibbonControl) '
Dim ConcatRange As Range
Dim cell As Range

Set ConcatRange = Intersect(Selection, Columns(ActiveCell.Column))

For Each cell In ConcatRange

Select Case Selection.Columns.Count
    Case 11: cell.Value = cell & " " & cell.Offset(0, 1) & " " & cell.Offset(0, 2) & " " & cell.Offset(0, 3) & " " & cell.Offset(0, 4) & " " & cell.Offset(0, 5) & " " & cell.Offset(0, 6) & " " & cell.Offset(0, 7) & " " & cell.Offset(0, 8) & " " & cell.Offset(0, 9) & " " & cell.Offset(0, 10)
    Case 10: cell.Value = cell & " " & cell.Offset(0, 1) & " " & cell.Offset(0, 2) & " " & cell.Offset(0, 3) & " " & cell.Offset(0, 4) & " " & cell.Offset(0, 5) & " " & cell.Offset(0, 6) & " " & cell.Offset(0, 7) & " " & cell.Offset(0, 8) & " " & cell.Offset(0, 9)
    Case 9: cell.Value = cell & " " & cell.Offset(0, 1) & " " & cell.Offset(0, 2) & " " & cell.Offset(0, 3) & " " & cell.Offset(0, 4) & " " & cell.Offset(0, 5) & " " & cell.Offset(0, 6) & " " & cell.Offset(0, 7) & " " & cell.Offset(0, 8)
    Case 8: cell.Value = cell & " " & cell.Offset(0, 1) & " " & cell.Offset(0, 2) & " " & cell.Offset(0, 3) & " " & cell.Offset(0, 4) & " " & cell.Offset(0, 5) & " " & cell.Offset(0, 6) & " " & cell.Offset(0, 7)
    Case 7: cell.Value = cell & " " & cell.Offset(0, 1) & " " & cell.Offset(0, 2) & " " & cell.Offset(0, 3) & " " & cell.Offset(0, 4) & " " & cell.Offset(0, 5) & " " & cell.Offset(0, 6)
    Case 6: cell.Value = cell & " " & cell.Offset(0, 1) & " " & cell.Offset(0, 2) & " " & cell.Offset(0, 3) & " " & cell.Offset(0, 4) & " " & cell.Offset(0, 5)
    Case 5: cell.Value = cell & " " & cell.Offset(0, 1) & " " & cell.Offset(0, 2) & " " & cell.Offset(0, 3) & " " & cell.Offset(0, 4)
    Case 4: cell.Value = cell & " " & cell.Offset(0, 1) & " " & cell.Offset(0, 2) & " " & cell.Offset(0, 3)
    Case 3: cell.Value = cell & " " & cell.Offset(0, 1) & " " & cell.Offset(0, 2)
    Case 2: cell.Value = cell & " " & cell.Offset(0, 1)

    Case Else: MsgBox ("Only 9 Columns can be concatenated with this function."): Exit Sub

End Select

Range(cell.Offset(0, 1), cell.Offset(0, Selection.Columns.Count - 1)).Clear

    Do Until Right(cell, 1) <> " "
        cell = Left(cell, Len(cell) - 1)
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
'This macro will import all the files in a given directory into a new workbook
'in Excel. The files will be imported based on the static filed sizes designated
'for all external files using the PeopleSoft "Upload Process". This is all files
'except for Time America files.

'This is a useful tool for validating files loads and reviewing time in the
'external time files.
Dim DirectoryPath As String
Dim FileName As String

'Allow the user to select the directory with the files
With Application.FileDialog(msoFileDialogFolderPicker)
        .InitialFileName = CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\"
        .Title = "Select the folder with external time files."
        .Show
        
    Select Case .SelectedItems.Count
        Case Is = 0: Exit Sub
        Case Is = 1: DirectoryPath = .SelectedItems(1) & "\"
    End Select
    
End With

Application.ScreenUpdating = False

'Create new sheet if current sheet is not empty
If ActiveSheet.UsedRange.Address <> "$A$1" Or Not IsEmpty(Range("A1")) Then
    Sheets.Add
    ActiveSheet.Name = "Time America Files"
End If

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

'Iterate through the files
FileName = Dir(DirectoryPath, vbReadOnly) ' + vbHidden)

    Do While FileName <> ""

        If UCase(Right(FileName, 4)) = ".TXT" Or UCase(Right(FileName, 4)) = ".DAT" Then
        
            'Import the files using delimination specific to the file format.
            With ActiveSheet.QueryTables.Add(Connection:= _
                "TEXT;" & DirectoryPath & FileName _
                , Destination:=Range(ActiveCell.Address))
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
                
                'List of field names and sizes that all external time files with this file format adhear to:
                
                '$emplid:11 $emplrcd:3 $reportdate:10 $trc:5 $hours:6 $amt:8 $profile:1 $business_unit:5 $deptid:10
                '$account:6 $product:6 $project_id:15 $business_unit_pc:5 $activity_id:15 $resource_type:5
                '$resource_cat:5 $resource_sub_cat:5
                    
                .TextFileColumnDataTypes = Array(1, 1, 3, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1)
                .TextFileFixedColumnWidths = Array(11, 3, 10, 5, 6, 8, 1, 5, 10, 6, 6, 15, 5, 15, 5, 5, 5)
                .TextFileTrailingMinusNumbers = True
                .Refresh BackgroundQuery:=False
            End With
            
        'Delete the data connection that pulled in the file data
        ActiveWorkbook.Connections(Left(FileName, Len(FileName) - 4)).Delete
        
        'Set the values of the cells in column A equal to the source filename.
        Range(ActiveCell.Offset(0, -1).Address, ActiveCell.End(xlDown).Offset(0, -1)) = Left(FileName, Len(FileName) - 4)
        
        'Change the active cell to be the next available cell.
        Range("B" & ActiveSheet.UsedRange.SpecialCells(xlLastCell).Row + 1).Select
        
        End If

FileName = Dir

Loop
    
'Do some formatting and move the worksheet to it's own workbook.
    Columns("D:D").NumberFormat = "m/d/yyyy"
    ActiveSheet.Name = "External Files PP" & RecentPP()
    Columns.AutoFit
    Range("A1").Select
    If Sheets.Count > 1 Then ActiveSheet.Move
    
Cells.EntireColumn.AutoFit
Rows(1).Font.Bold = True

'TODO: Add a pivot table to see the total hours for each file.

Application.ScreenUpdating = True

End Sub

Sub ImportTimeAmericaFiles()
'This macro will import all the Time America files in a given directory into the
'current Excel workbook. The files will be deliminated by commas.
'This is a useful tool for validating files loads and reviewing time in the external time files.

Dim DirectoryPath As String
Dim FileName As String

'Allow the user to select the directory with the files
With Application.FileDialog(msoFileDialogFolderPicker)
        .InitialFileName = "\\L12239\CXFUSR\Appl\HR800\PS\Temp\GSC\"
        .Title = "Select the folder with external time files."
        .Show
        
    Select Case .SelectedItems.Count
        Case Is = 0: Exit Sub
        Case Is = 1: DirectoryPath = .SelectedItems(1) & "\"
    End Select
    
End With

Application.ScreenUpdating = False

'Create new sheet if current sheet is not empty
If ActiveSheet.UsedRange.Address <> "$A$1" Or Not IsEmpty(Range("A1")) Then
    Sheets.Add
    ActiveSheet.Name = "Time America Files"
End If

'Set headers
    Range("A1").FormulaR1C1 = "EmpID"
    Range("B1").FormulaR1C1 = "TRC"
    Range("C1").FormulaR1C1 = "Hours"
    Range("D1").FormulaR1C1 = "Reported Date"
    Range("E1").FormulaR1C1 = "File Name"

Range("A2").Select

'Iterate through the files
FileName = Dir(DirectoryPath, vbReadOnly) ' + vbHidden)

    Do While FileName <> ""

        If UCase(Right(FileName, 4)) = ".TXT" Then
        
            'Import the files using delimination specific to the file format.
            With ActiveSheet.QueryTables.Add(Connection:= _
                "TEXT;" & DirectoryPath & FileName _
                , Destination:=Range(ActiveCell.Address))
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
                .TextFileParseType = xlDelimited
                .TextFileTextQualifier = xlTextQualifierDoubleQuote
                .TextFileConsecutiveDelimiter = False
                .TextFileTabDelimiter = False
                .TextFileSemicolonDelimiter = False
                .TextFileCommaDelimiter = True
                .TextFileSpaceDelimiter = False
                .TextFileColumnDataTypes = Array(9, 1, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 1, 1, 5, 9, 9, 9, 9, 9, _
                9, 9, 9, 9)
                .TextFileTrailingMinusNumbers = True
                .Refresh BackgroundQuery:=False
            End With
        
        'Delete the data connection that pulled in the file data
        ActiveWorkbook.Connections(Left(FileName, Len(FileName) - 4)).Delete
        
        'Set the values of the cells in column A equal to the source filename.
        Range(ActiveCell.Offset(0, 4).Address, ActiveCell.End(xlDown).Offset(0, 4)) = Left(FileName, Len(FileName) - 4)
        
        'Change the active cell to be the next available cell.
        Range("A" & ActiveSheet.UsedRange.SpecialCells(xlLastCell).Row + 1).Select
        
        End If

FileName = Dir

Loop

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
