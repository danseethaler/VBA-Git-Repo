Attribute VB_Name = "TimeAmericaProcessing"
Option Explicit

Sub AppendFiles(control As IRibbonControl)
'This macro will aggregate all the files with extension .txt in the selected folder
'into a new file. This is used to aggregate the external time files from Time America
'into a single file that can be uploaded a single time rather than uploading dozens
'of seperate files. All original files remain unaltered.

Dim SourceNum As Integer
Dim DestNum As Integer, FileCount As Long
Dim Temp As String, directoryPath As String
Dim fileName As String, cell As Range
Dim AggFile As String
Dim FMACounter As Integer

'Determine the filename for the new file.
AggFile = InputBox("Please provide a name for the aggregated file.", "File Name") & ".txt"

'If the user did not type anything in for a filename then the macro will terminate
If AggFile = ".txt" Then Exit Sub

'Ask the user for the folder that contains the files to merge
    With Application.FileDialog(msoFileDialogFolderPicker)
        .InitialFileName = "\\L12239\CXFUSR\Appl\HR800\PS\Temp\GSC\"
        .Title = "Select the folder with texts files to merge."
        .Show
    
    'If no folder is selected then terminate the macro.
    'Otherwise set the DirectoryPath variable equal to the path of the selected folder.
    Select Case .SelectedItems.Count
        Case Is = 0: Exit Sub
        Case Is = 1: directoryPath = .SelectedItems(1) & "\"
    End Select
    
    End With

fileName = Dir(directoryPath, vbReadOnly) ' + vbHidden)
    
Application.ScreenUpdating = False
    
Do While fileName <> ""
    Application.StatusBar = FileCount & " Files Complete..."
    If fileName <> AggFile And InStr(UCase(fileName), ".TXT") > 0 Then
      ' Open the destination text file.
      DestNum = FreeFile()
      Open directoryPath & AggFile For Append As DestNum

      ' Open the source text file.
      SourceNum = FreeFile()
      Open directoryPath & fileName For Input As SourceNum

      ' Read each line of the source file and append it to the destination file.
      Do While Not EOF(SourceNum)
           Line Input #SourceNum, Temp
           
           'Add the source line to the destination file.
           Print #DestNum, Temp
           
        Loop
        
        'Get the next file in the directory an increase the number of files processed by one.
          fileName = Dir
          FileCount = FileCount + 1
        
      'Close the files we've been using. The new file will be opened in the next loop.
      Close #DestNum
      Close #SourceNum
    
    'If the file does not have a .txt extension then move onto the next file.
    Else: fileName = Dir

    End If
        
    Loop

'Reset the status bar and begin updating the screen again.
Application.ScreenUpdating = True
Application.StatusBar = False
    
End Sub

Sub ImportTimeAmericaFiles(control As IRibbonControl)
'This macro will import all the Time America files in a given directory into the
'current Excel workbook. The files will be deliminated by commas.
'This is a useful tool for validating files loads and reviewing time in the external time files.

Dim directoryPath As String
Dim fileName As String

directoryPath = "\\L12239\cxfusr\Appl\HR800\PS\Temp\GSC\"

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
fileName = Dir(directoryPath, vbReadOnly) ' + vbHidden)

    Do While fileName <> ""

        If UCase(Right(fileName, 4)) = ".TXT" Then
        
            'Import the files using delimination specific to the file format.
            With ActiveSheet.QueryTables.Add(Connection:= _
                "TEXT;" & directoryPath & fileName _
                , Destination:=Range(ActiveCell.Address))
                .Name = Left(fileName, Len(fileName) - 4)
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
        ActiveWorkbook.Connections(Left(fileName, Len(fileName) - 4)).Delete
        
        'Set the values of the cells in column A equal to the source filename.
        Range(ActiveCell.Offset(0, 4).Address, ActiveCell.End(xlDown).Offset(0, 4)) = Left(fileName, Len(fileName) - 4)
        
        'Change the active cell to be the next available cell.
        Range("A" & ActiveSheet.UsedRange.SpecialCells(xlLastCell).Row + 1).Select
        
        End If

fileName = Dir

Loop

Application.ScreenUpdating = True

End Sub

Sub ImportExternalTimeFile(control As IRibbonControl)
'This macro will import all the files in a given directory into a new workbook
'in Excel. The files will be imported based on the static filed sizes designated
'for all external files using the PeopleSoft "Upload Process". This is all files
'except for Time America files.

'This is a useful tool for validating files loads and reviewing time in the
'external time files.
Dim directoryPath As String
Dim fileName As String

'Allow the user to select the directory with the files
With Application.FileDialog(msoFileDialogFolderPicker)
        .InitialFileName = CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\"
        .Title = "Select the folder with external time files."
        .Show
        
    Select Case .SelectedItems.Count
        Case Is = 0: Exit Sub
        Case Is = 1: directoryPath = .SelectedItems(1) & "\"
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
fileName = Dir(directoryPath, vbReadOnly) ' + vbHidden)

    Do While fileName <> ""

        If UCase(Right(fileName, 4)) = ".TXT" Or UCase(Right(fileName, 4)) = ".DAT" Then
        
            'Import the files using delimination specific to the file format.
            With ActiveSheet.QueryTables.Add(Connection:= _
                "TEXT;" & directoryPath & fileName _
                , Destination:=Range(ActiveCell.Address))
                .Name = Left(fileName, Len(fileName) - 4)
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
        ActiveWorkbook.Connections(Left(fileName, Len(fileName) - 4)).Delete
        
        'Set the values of the cells in column A equal to the source filename.
        Range(ActiveCell.Offset(0, -1).Address, ActiveCell.End(xlDown).Offset(0, -1)) = Left(fileName, Len(fileName) - 4)
        
        'Change the active cell to be the next available cell.
        Range("B" & ActiveSheet.UsedRange.SpecialCells(xlLastCell).Row + 1).Select
        
        End If

fileName = Dir

Loop
    
'Do some formatting and move the worksheet to it's own workbook.
    Columns("D:D").NumberFormat = "m/d/yyyy"
    ActiveSheet.Name = "External Files PP" & RecentPPforTA()
    Columns.AutoFit
    Range("A1").Select
    If Sheets.Count > 1 Then ActiveSheet.Move
    
Cells.EntireColumn.AutoFit
Rows(1).Font.Bold = True

'TODO: Add a pivot table to see the total hours for each file.

Application.ScreenUpdating = True

End Sub

Sub PrepareTAErrorReport(control As IRibbonControl)
'This macro prepares the TA Error Report to be sent.

Dim lastRow As Integer
Dim i As Integer

Dim parameter As IRibbonControl

If ActiveSheet.Name <> "Errors" Then
    MsgBox "Please make sure to be on the Errors tab of the DI Email Details workbook before continuing."
    Exit Sub
End If

'Clear sheet 3 if there is any content.
If Sheets(3).UsedRange.Address <> "$A$1" Or Not IsEmpty(Sheets(3).Range("A1")) Then
    Sheets(3).UsedRange.ClearContents
End If

lastRow = ActiveSheet.UsedRange.SpecialCells(xlLastCell).Row

    Range("A2:A" & lastRow).TextToColumns Destination:=Range("A2"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=True, Tab:=False, _
        Semicolon:=False, Comma:=False, Space:=True, Other:=False, FieldInfo _
        :=Array(Array(1, 1), Array(2, 1), Array(3, 1), Array(4, 1), Array(5, 1), Array(6, 1), _
        Array(7, 1), Array(8, 1), Array(9, 1), Array(10, 1), Array(11, 1), Array(12, 1)), _
        TrailingMinusNumbers:=True
    Range("A2:A" & lastRow).ClearContents
    Columns("C:C").Delete Shift:=xlToLeft
    Range("F2:L" & lastRow).Select
    
    Call ConcatenateDelimitedText(parameter)
    
    Range("A1").Select
    Sheets("File Details").Select
    
    Call ImportTimeAmericaFiles(parameter)
    
    Range("A1").Select
    
    Call ColumnsAutofit(parameter)
    
    Sheets("Errors").Select
    Range("A2:A" & lastRow).FormulaR1C1 = "=VLOOKUP(RC[1],'File Details'!C:C[4],5,FALSE)"
    Range("A2:A" & lastRow).Copy
    Range("A2:A" & lastRow).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    For i = Range("A1").End(xlDown).Row To 2 Step -1
        'Delete any Multiple Employee Record errors - these are not an issue that the DI stores need to fix.
        If Range("A" & i).Offset(0, 5) = "Multiple Active Empl_Rcds Exists. Using latest" Then
            Range("A" & i).EntireRow.Delete
        End If
        
    Next i
        
    Range("B2:B" & lastRow).Select
    
    Call ConvertEmpIDToText(parameter)
    
    Range("A1").Select
    
    ActiveWorkbook.Save

End Sub

Sub TimeAmericaErrorReport(control As IRibbonControl)
'This macro is used to generate an email which notifies the DI stores of
'errors generated when uploading the Time America file from their store.

    Dim cell As Range
    Dim Outlook As Outlook.Application
    Dim MyItem As Outlook.MailItem
    Dim EmailTemplate As String
    Dim Stores As String
    Dim i As Integer
    
    'Make sure Outlook is open.
    On Error Resume Next
    Set Outlook = GetObject(, "Outlook.Application")
    On Error GoTo 0
    
    'If Outlook is not open then open it.
    If Outlook Is Nothing Then
        Set Outlook = CreateObject("Outlook.Application")
    End If
    
    'Set the Email Template string variable equal to the directory of the Outlook email template
    EmailTemplate = "\\CHQPVUN0066\FINUSR\SHARED\FIN_PYRL\2_Payroll Time & Labor Absence Management\Desk Manual (Information)\Time America Processing Files\TA100 Uploads Template.oft"
    
    'Remove any formulas in column A.
    Range("A:A").Value = Range("A:A").Value

    'Iterate through the error stores and generate a list of store names with errors on their file upload.
    If Not IsEmpty(Range("A2")) Then
        For i = Range("A1").End(xlDown).Row To 2 Step -1
            'EmpID 353605 appears on this error report even though the job data is accruate.
            If Range("A" & i).Offset(0, 1) = "353605" Then
                Range("A" & i).EntireRow.Delete
            End If
            
            If Not IsEmpty(Range("A" & i)) And InStr(1, Stores, Range("A" & i)) = 0 Then
                Stores = Stores & Range("A" & i).Value & "<br>"
            End If

        Next i
    End If
    
    'Range("A2") should have the storename associated with that error in it.
    'If not the macro will exit.
    If IsEmpty(Range("A2")) Then
        MsgBox "There is not a Store Name in cell A2. Please update the error report and rerun the macro."
        Exit Sub
    End If
    
    'Remove any astrisks from the worksheet (they come in from the error page).
    Cells.Replace What:="~*", Replacement:="", LookAt:=xlPart, SearchOrder _
    :=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
    
    Range("A:A").HorizontalAlignment = xlLeft
    
    'Save the workbook.
    ActiveWorkbook.Save
    
    'Copy the worksheet with the errors
    Application.ActiveSheet.Copy
    ActiveWorkbook.SaveAs (CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\PP" & RecentPPforTA() & " Load Errors")
    
    'This line creates the Outlook mail object and assigns it to the designated template.
    Set MyItem = Outlook.CreateItemFromTemplate(EmailTemplate)
    
    'This section manipulates several of the properties of the email template to insert
    'the information from the worksheet into the email.
    With MyItem
        .Attachments.Add ActiveWorkbook.FullName
        .To = "DL-WEL-DIStaff"
        .CC = "danseethaler@ldschurch.org;awilkins@ldschurch.org;DL-GSC-PrcSvc-PR-EmployeeData@ldschurch.org;" & _
                "ashkw@ldschurch.org;ThuesonJJ@ldschurch.org;wrigleyjc@ldschurch.org;WiddisonKL@ldschurch.org;" & _
                "MoorePJ@ldschurch.org;GriffinHP@ldschurch.org;WarrinerTS@ldschurch.org"
        .BCC = ""
        .Subject = .Subject & RecentPPforTA()
        .HTMLBody = Replace(MyItem.HTMLBody, "#Stores", Stores)
        .Display
    End With
    
    'Close and delete the workbook.
    Application.ActiveWorkbook.Close
    Kill (CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\PP" & RecentPPforTA() & " Load Errors.xlsx")

'Remind the operator to run the Employee Time by TRC report to catch further upload errors.
MsgBox ("Please run the employee time by TRC report and make corrections to errors.")

End Sub

Sub EmailMissingDIStores(control As IRibbonControl)
'This macro generates a dictionary object that is used to validate all DI stores
'have submitted their Time America file and it is ready to process.
'The output of this macro is an email to the stores who have not yet sent us their
'Time America file.
'The email addresses listed in the dictionary below need to be manually updated as needed.

    Dim ToList As String
    Dim directoryPath As String
    Dim fileName As String
    Dim missingStoresList As String
    Dim cell As Range
    Dim i As Integer
    Dim Outlook As Outlook.Application
    Dim MyItem As Outlook.MailItem
    
    Dim clipboard As MSForms.DataObject
    Set clipboard = New MSForms.DataObject
    
    Dim Stores As Dictionary
    Set Stores = New Dictionary
    Dim strKey As Variant
    
    directoryPath = "\\L12239\CXFUSR\Appl\HR800\PS\Temp\GSC\"

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
    .Add Key:="Fontana", Item:="pcampbell@ldschurch.org;MasseyDJ@ldschurch.org"
    .Add Key:="Downtown SLC", Item:="LoseeWe@ldschurch.org;SorensenJE@ldschurch.org"
    .Add Key:="Federal Way", Item:="Barbara.Hellickson@ldschurch.org;ClementGL@ldschurch.org"
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
    .Add Key:="Price", Item:="colleen.byrge@ldschurch.org;matthew.kemp@ldschurch.org"
    .Add Key:="Provo", Item:="NelsonP@ldschurch.org;OlsonWa@ldschurch.org"
    .Add Key:="Rexburg", Item:="tracy.smith@ldschurch.org;GlissmeyerKG@ldschurch.org"
    .Add Key:="Richfield", Item:="alicia.murray@ldschurch.org;BaroneMa@ldschurch.org"
    .Add Key:="Sacramento", Item:="btourtillott@ldschurch.org;ThomasTD@ldschurch.org"
    .Add Key:="Sandy", Item:="krista.loiacono@ldschurch.org;MontalboMA@ldschurch.org"
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

    fileName = Dir(directoryPath)
    Do While fileName <> ""
        'Remove the Store Name/File Name from the dictionary
        If Stores.Exists(Left(fileName, InStrRev(fileName, ".") - 1)) Then
            Stores.Remove Left(fileName, InStrRev(fileName, ".") - 1)
        Else
            'If the filename does not match a dictionary member a message is generated to the user.
            MsgBox "Filename " & fileName & " does not match a member of the stores dictionary."
        End If
        
        fileName = Dir
    Loop

    'Iterate through the remaining dictionary keys and add the associated email addresses to the ToList string variable.
    For Each strKey In Stores.Keys()
        ToList = ToList & Stores(strKey) & ";"
        missingStoresList = missingStoresList & strKey & vbNewLine
    Next
    
    'If all dictionary keys have been removed then we can safely say all files are ready to load.
    If Stores.Count = 0 Then
            MsgBox "All DI files have been received."
            Exit Sub
        Else
            clipboard.SetText missingStoresList
            clipboard.PutInClipboard
    End If

    
    'Make sure Outlook is open.
    On Error Resume Next
    Set Outlook = GetObject(, "Outlook.Application")
    On Error GoTo 0
    
    'If Outlook is not open then open it.
    If Outlook Is Nothing Then
        Set Outlook = CreateObject("Outlook.Application")
    End If
    
    'Create the email template from the specified directory path
    Set MyItem = Outlook.CreateItemFromTemplate("\\chqpvun0066\finusr\SHARED\FIN_PYRL\2_Payroll Time & Labor Absence Management\Desk Manual (Information)\Time America Processing Files\DI TA100 Missing.oft")
    
    'Update the email with information from the macro and display the email.
    With MyItem
        .To = ToList
        .CC = "danseethaler@ldschurch.org;awilkins@ldschurch.org;"
        .BCC = ""
        .Subject = "Missing the PP" & RecentPPforTA() & " Time America TXT File"
        .HTMLBody = Replace(.HTMLBody, "#PP", RecentPPforTA())
        .Display
    End With

End Sub

Public Function RecentPPforTA() As Integer
'This function returns the most recently completed pay period number as an integer.

Dim PP01 As Date
Dim lastPPEndDate As Date

'Determine the most recent PP end date based on a static date in the past
lastPPEndDate = Date - (Date - CDate("8/15/2014")) Mod 14

'Determine the PPEnd date for PP01 of the current year
PP01 = CDate("1/3/2014")
Do Until Year(PP01 + 7) = Year(Date)
    PP01 = PP01 + 14
Loop

'If the first PP end date this year is after the current date then we assume the most recent PP
'is pay period 26.
If PP01 >= Date Then
    RecentPPforTA = 26
    Exit Function
End If

'Determine which pay period just ended.
RecentPPforTA = 1
Do Until lastPPEndDate = PP01
    PP01 = PP01 + 14
    RecentPPforTA = RecentPPforTA + 1
Loop

End Function


