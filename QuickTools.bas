Attribute VB_Name = "QuickTools"
Option Explicit
Dim WorkRange As Range

Sub ConvertEmpIDToText(control As IRibbonControl)
Dim Cell As Range
Dim EmpIDCount As Integer

Application.ScreenUpdating = False

On Error Resume Next

    If Selection.Cells.Count = 1 Then
        If Not IsEmpty(ActiveCell) And Not ActiveCell.HasFormula Then
            
            Set WorkRange = Selection
        
        Else

            MsgBox "Please select a range with constants", vbInformation
            Application.ScreenUpdating = True
            Exit Sub
        
        End If
    
    ElseIf Selection.Cells.Count > 1 Then

        Set WorkRange = Selection.SpecialCells(xlCellTypeConstants)
        Set WorkRange = Intersect(WorkRange, WorkRange.Parent.UsedRange)
    
    End If
    
    If Err = 1004 Or Err = 94 Or Err = 91 Then
        MsgBox "There are no constants in your selection", vbInformation
        Application.ScreenUpdating = True
        Exit Sub
    ElseIf Err <> 0 Then
        MsgBox "Error " & Err & ": " & Error(Err.Number)
        Application.ScreenUpdating = True
        Exit Sub
    End If
    
On Error GoTo 0
    
    If WorkRange Is Nothing Then
        MsgBox "Please select a valid range."
        Application.ScreenUpdating = True
        Exit Sub
    End If

For Each Cell In WorkRange

    Do Until Right(Cell, 1) <> " "
        Cell = Left(Cell, Len(Cell) - 1)
    Loop

If IsNumeric(Cell) Then

    Cell.NumberFormat = "@"
    
        If Len(Cell) = 1 Then
            Cell.Value = "00000" & Cell.Value
        ElseIf Len(Cell) = 2 Then
            Cell.Value = "0000" & Cell.Value
        ElseIf Len(Cell) = 3 Then
            Cell.Value = "000" & Cell.Value
        ElseIf Len(Cell) = 4 Then
            Cell.Value = "00" & Cell.Value
        ElseIf Len(Cell) = 5 Then
            Cell.Value = "0" & Cell.Value
        End If
        
        If Len(Cell) = 6 Then
            Cell.Value = "'" & Cell.Value
            'EmpIDCount = EmpIDCount + 1
        End If

End If
        
Next Cell

Call UsageLog("EmpID to Text")

Application.ScreenUpdating = True

'MsgBox (EmpIDCount & " EmpID(s) converted to text.")

End Sub

Sub ConvertSSNToText(control As IRibbonControl)
Dim Cell As Range
Dim WorkRange As Range
Dim SSNCount As Integer

On Error Resume Next

    If Selection.Cells.Count = 1 Then
        If Not IsEmpty(ActiveCell) And Not ActiveCell.HasFormula Then
            
            Set WorkRange = Selection
        
        Else

            MsgBox "Please select a range with constants", vbInformation
            Exit Sub
        
        End If
    
    ElseIf Selection.Cells.Count > 1 Then

        Set WorkRange = Selection.SpecialCells(xlCellTypeConstants)
        Set WorkRange = Intersect(WorkRange, WorkRange.Parent.UsedRange)
    
    End If
    
    If Err = 1004 Or Err = 94 Or Err = 91 Then
        MsgBox "There are no constants in your selection", vbInformation
        Exit Sub
    ElseIf Err <> 0 Then
        MsgBox "Error " & Err & ": " & Error(Err.Number)
        Exit Sub
    End If
    
On Error GoTo 0
    
    If WorkRange Is Nothing Then
        MsgBox "Please select a valid range."
        Exit Sub
    End If

WorkRange.Replace What:="-", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False

For Each Cell In WorkRange

If IsNumeric(Cell) Then

    Cell.NumberFormat = "@"
    
        If Len(Cell) = 6 Then
            Cell.Value = "000" & Cell.Value
        ElseIf Len(Cell) = 7 Then
            Cell.Value = "00" & Cell.Value
        ElseIf Len(Cell) = 8 Then
            Cell.Value = "0" & Cell.Value
        End If

        If Len(Cell) = 9 Then
            Cell.Value = "'" & Cell.Value
            SSNCount = SSNCount + 1
        End If
End If
        
Next Cell

Call UsageLog("SSN to Text")

'MsgBox (SSNCount & " SSN(s) converted to text.")

End Sub

Sub PSTable(control As IRibbonControl)
Dim TableName As String
Dim ColumnCount As Integer
Dim ActiveColumn As Integer
Dim Cell As Range
Dim DConnection As Integer

Application.ScreenUpdating = False

If Range("A2").Interior.Color = 12632256 Then

On Error Resume Next

    If UCase(Left(ActiveSheet.Name, 5)) = "SHEET" Then
    ActiveSheet.Name = Left(Range("A1"), 28 - Len(Range("B1"))) & " (" & Replace(Range("B1"), " ", "") & ")"
    End If

On Error GoTo 0

    Rows("1:1").Delete Shift:=xlUp
    
End If

If ActiveSheet.ListObjects.Count = 1 Then ActiveSheet.ListObjects(1).Unlist

    With Range("A1").CurrentRegion
        .Font.Name = "Calibri"
        .Font.Size = 11
    End With

If Len(Range("B1")) > 0 And Range("A1").Interior.PatternColor = 0 Then
    With Range("A1:" & Cells(1, Range("A1").End(xlToRight).Column).Address)
        .Font.Bold = False
        .ClearFormats
    End With
    With Range("A1:" & Cells(1, Range("A1").End(xlToRight).Column).Address)
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Range("A1:" & Cells(1, Range("A1").End(xlToRight).Column).Address).Font
        .Name = "Calibri"
        .FontStyle = "Regular"
        .Size = 11
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        '.Underline = xlUnderlineStyleSingleAccounting
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    
    With Range("A1:" & Cells(1, Range("A1").End(xlToRight).Column).Address).Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.249977111117893
        .PatternTintAndShade = 0
    End With

For Each Cell In Range("A1:" & Cells(1, Range("A1").End(xlToRight).Column).Address)
        
    Cell.HorizontalAlignment = Cell.Offset(1, 0).HorizontalAlignment
    
    If IsNumeric(Cell.Offset(1, 0)) Or IsDate(Cell.Offset(1, 0)) Then Cell.HorizontalAlignment = xlRight
    If Cell.Offset(1, 0).NumberFormat = "@" Then Cell.HorizontalAlignment = xlLeft

Cell.VerticalAlignment = xlCenter

Next Cell

End If

If Not ActiveWindow.DisplayGridlines Then

ColumnCount = Range("A1:" & ActiveSheet.UsedRange.SpecialCells(xlCellTypeLastCell).Address).Columns.Count

On Error Resume Next

For ActiveColumn = 1 To ColumnCount Step 1

    If Columns(ActiveColumn).ColumnWidth < 60 Then
        Columns(ActiveColumn).ColumnWidth = 60
    End If

Next

End If
    
    ActiveWindow.DisplayGridlines = True
    
    If ActiveSheet.ListObjects.Count < 1 Then
        If ActiveWorkbook.Connections.Count > 0 Then
        For DConnection = 1 To ActiveWorkbook.Connections.Count
        ActiveWorkbook.Connections(1).Delete
        Next DConnection
    End If
    
'    If ActiveWorkbook.Names.Count > 1 Then
'        For DConnection = 1 To ActiveWorkbook.Names.Count
'        ActiveWorkbook.Names(1).Delete
'        Next DConnection
'    ElseIf ActiveWorkbook.Names.Count = 1 Then
'        ActiveWorkbook.Names(1).Delete
'    End If
    
    ActiveSheet.ListObjects.Add(xlSrcRange, Range("A1").CurrentRegion, , xlYes).Name _
        = "Table" & ActiveSheet.Index
    Range("A1").CurrentRegion.Select
    ActiveSheet.ListObjects(1).TableStyle = ""
    
    Selection.AutoFilter
    
    End If
    
    With Selection
        .Borders(xlDiagonalDown).LineStyle = xlNone
        .Borders(xlDiagonalUp).LineStyle = xlNone
        .Borders(xlEdgeLeft).LineStyle = xlNone
        .Borders(xlEdgeTop).LineStyle = xlNone
        .Borders(xlEdgeBottom).LineStyle = xlNone
        .Borders(xlEdgeRight).LineStyle = xlNone
        .Borders(xlInsideVertical).LineStyle = xlNone
        .Borders(xlInsideHorizontal).LineStyle = xlNone
    End With
    
    Call ColumnsAutofitCall

Call UsageLog("Format PeopleSoft Query")

Range("A1").Select

Application.ScreenUpdating = True

End Sub

Sub ListCriteria(control As IRibbonControl)

Dim Cell As Range
Dim Message As String
Dim Message1 As String
Dim List As String
Dim clipboard As MSForms.DataObject
Set clipboard = New MSForms.DataObject

If TypeName(Selection) <> "Range" Then Exit Sub

Application.ScreenUpdating = False

If Selection.Cells.Count > 1 Then
MsgBox "Psst... I'm going to select just the first cell of your selection and then continue." & vbNewLine & vbNewLine & _
    "Then I will automatically select as many cells as will fit in a PeopleSoft list criteria." & vbNewLine & vbNewLine & _
    "I'll also skip hidden cells and ignore duplicate values." & vbNewLine & vbNewLine & _
    "Finally, I'll activate the next unique value in the column so you can simply run the macro again " & _
    "to finish selecting the unique values. Pretty cool huh?"
Range(Cells(Selection.Row, Selection.Column).Address(0, 0)).Select

End If

    If Selection.Cells.Count = 1 Then
        If Not IsEmpty(ActiveCell) And Not ActiveCell.HasFormula Then
        
        Do Until Len(Message) - 3 > 245
            If ActiveCell.EntireRow.Hidden <> True Then
            If InStr(1, Message, ActiveCell) = 0 And Not IsEmpty(ActiveCell) Then
            
                Do Until Right(ActiveCell, 1) <> " "
                ActiveCell = Left(ActiveCell, Len(ActiveCell) - 1)
                Loop
                
                Message1 = Message & ActiveCell.Value & "','"
                
                If Len(Message1) > 255 Then GoTo MaxMessage
                
                Message = Message & ActiveCell.Value & "','"
            End If
            End If
        ActiveCell.Offset(1, 0).Select
        
        If IsEmpty(ActiveCell) Then
            Selection.Offset(-1, 0).End(xlUp).Select
            Exit Do
        End If
        
        Loop
        
        End If
    
    ElseIf Selection.Cells.Count > 1 Then

        Set WorkRange = Selection.SpecialCells(xlCellTypeVisible)
        Set WorkRange = WorkRange.SpecialCells(xlCellTypeConstants)
        Set WorkRange = Intersect(WorkRange, WorkRange.Parent.UsedRange)
    
    If Err = 1004 Or Err = 94 Or Err = 91 Then
        MsgBox "There are no constants in your selection", vbInformation
        Exit Sub
        
    ElseIf Err = 6 Then
    
        Set WorkRange = Selection.SpecialCells(xlCellTypeConstants)
        Set WorkRange = Intersect(WorkRange, WorkRange.Parent.UsedRange)
    
    
    ElseIf Err <> 0 Then
        MsgBox "Error " & Err & ": " & Error(Err.Number), vbCritical
        Exit Sub
    End If

For Each Cell In WorkRange

If Cell.Rows(Cell.Row).Hidden = True Then GoTo SkipCell

If Not IsEmpty(Cell) And InStr(1, Message, Cell) = 0 Then
    Message = Message & Cell.Value & "','"
End If

SkipCell:

Next Cell

End If

MaxMessage:

Application.ScreenUpdating = True

Selection.Select

If Selection.Row <> 1 Then

Do Until Selection <> Selection.Offset(-1, 0)
Selection.Offset(1, 0).Select
Loop

End If

If IsEmpty(Selection) Then Selection.Offset(-1, 0).End(xlUp).Select

If Len(Left(Message, Len(Message) - 3)) > 255 Then
MsgBox ("Each list member in PeopleSoft can contain a maximum of 28 EmplIDs." & vbNewLine & vbNewLine & _
    "Please shrink your selection size and then add multiple selections to the PeopleSoft list members."), vbCritical
Exit Sub
End If

Message = Left(Message, Len(Message) - 3)

clipboard.SetText Message
clipboard.PutInClipboard

MsgBox ("This list has been copied to your clipboard." & vbNewLine & vbNewLine & _
        "You can paste it into the PeopleSoft list members to add.")
        
Call UsageLog("List Criteria")

End Sub

Sub ReopenUnsavedWorkbook(control As IRibbonControl)
    Dim ConfirmSub
    Dim openWorkbook As String
    
    'ConfirmSub = MsgBox("Do you want to close this workbook without saving?", vbYesNo + vbDefaultButton2 + vbCritical, "Confirm Close")
    
'    If ConfirmSub = vbNo Then
'        Exit Sub
'    End If

    openWorkbook = ActiveWorkbook.FullName

    Application.DisplayAlerts = False
        ActiveWorkbook.Close (False)
    Application.DisplayAlerts = True
    
    Application.Workbooks.Open (openWorkbook)
    
Call UsageLog("ReOpen No Save")

End Sub

Sub KillWorkbook(control As IRibbonControl)

    Application.DisplayAlerts = False
        ActiveWorkbook.Close (True)
    Application.DisplayAlerts = True
    
If Workbooks.Count = 0 Then Application.Quit

End Sub

Private Sub ReverseName(control As IRibbonControl)

Dim FirstName As String, LastName As String
Dim Cell As Range
Dim WorkRange As Range
Dim NameCounter As Integer

On Error Resume Next

    If Selection.Cells.Count = 1 Then
        If Not IsEmpty(ActiveCell) And Not ActiveCell.HasFormula Then
            
            Set WorkRange = Selection
        
        Else

            MsgBox "Please select a range with names in 'Last, First' format.", vbInformation
            Exit Sub
        
        End If
    
    ElseIf Selection.Cells.Count > 1 Then

        Set WorkRange = Selection.SpecialCells(xlCellTypeConstants)
        Set WorkRange = Intersect(WorkRange, WorkRange.Parent.UsedRange)
    
    End If
    
    If Err = 1004 Or Err = 94 Or Err = 91 Then
        MsgBox "There are no constants in your selection", vbInformation
        Exit Sub
    
    ElseIf Err = 6 Then
    
        Set WorkRange = Selection.SpecialCells(xlCellTypeConstants)
        Set WorkRange = Intersect(WorkRange, WorkRange.Parent.UsedRange)
    
    ElseIf Err <> 0 Then
        MsgBox "Error " & Err & ": " & Error(Err.Number), vbCritical
        Exit Sub
    End If

ErrorHandler:

On Error GoTo 0


Dim NameValue As Boolean
For Each Cell In WorkRange

    
    NameValue = Cell Like "[$,;,:]"
    If NameValue <> True Then
        
        If InStr(Cell, ",") Then
            
            LastName = Left(Cell, InStr(Cell, ",") - 1)
            FirstName = Right(Cell, Len(Cell) - InStr(Cell, ","))
        
            If Left(FirstName, 1) = " " Then
                FirstName = Right(FirstName, Len(FirstName) - 1)
            End If
            
            Cell = WorksheetFunction.Proper(FirstName & " " & LastName)

            Cell = Left(Cell, InStr(Cell, " ") - 1) & " " & Right(Cell, Len(Cell) - InStrRev(Cell, " "))
            
            NameCounter = NameCounter + 1
        
        End If
        
        If InStr(Cell, "  ") Then
            
            LastName = Left(Cell, InStr(Cell, "  ") - 1)
            FirstName = Right(Cell, Len(Cell) - InStr(Cell, "  "))
        
            If Left(FirstName, 1) = " " Then
                FirstName = Right(FirstName, Len(FirstName) - 1)
            End If
            
            Cell = WorksheetFunction.Proper(FirstName & " " & LastName)
            
            Cell = Left(Cell, InStr(Cell, " ") - 1) & " " & Right(Cell, Len(Cell) - InStrRev(Cell, " "))
            
            NameCounter = NameCounter + 1
        
        End If
        
        Do Until Left(Cell, 1) <> " "
            Cell = Right(Cell, Len(Cell) - 1)
        Loop
    
    End If

Next Cell
Call UsageLog("Reverse Name")

'MsgBox NameCounter & " name(s) were reformatted to 'First Last' format."

End Sub

Sub CopyWorkSheet(control As IRibbonControl)

If IsEmpty(ActiveSheet) Then Exit Sub

ActiveSheet.Copy After:=ActiveSheet

Call UsageLog("Copy Worksheet")

End Sub

Sub FormulastoValues(control As IRibbonControl)
Dim WorkRange As Range
Dim intArea As Integer
Dim rngCell As Range
'Dim continue As String
'
'continue = MsgBox("Would you like to convert all formulas in your selection to values?", _
'            vbYesNo, "Convert Formulas to Values")
'
'    If continue = vbNo Then Exit Sub
    
    Set WorkRange = Selection
    Set WorkRange = Intersect(WorkRange, WorkRange.Parent.UsedRange)

On Error Resume Next

    If Not WorkRange.Cells.Count = 1 Then
        Set WorkRange = WorkRange.SpecialCells(xlCellTypeFormulas)
    End If
    
    If Err = 1004 Or Err = 94 Then
        MsgBox "There are no formulas in your selection", vbInformation
        Exit Sub
    ElseIf Err <> 0 Then
        MsgBox "Error " & Err & ": " & Error(Err.Number), vbCritical
        Exit Sub
    End If
    
On Error GoTo 0
    
    If WorkRange Is Nothing Then
        MsgBox "Please select a valid range.", vbCritical
        Exit Sub
    End If

    If WorkRange.Cells.Count = 1 And IsEmpty(Selection.Cells(1, 1)) Then
        MsgBox "Please select a range with values", vbInformation
        Exit Sub
    End If

On Error Resume Next

    If WorkRange.Areas.Count > 1 Then
    
        For intArea = 1 To WorkRange.Areas.Count
            With WorkRange.Areas(intArea)
                'Debug.Print WorkRange.Areas(intArea).Address
               ' .NumberFormat = "General"
                .Value = .Value
            End With
        Next
        
    ElseIf WorkRange.Areas.Count = 1 Then
    
            With WorkRange
                .NumberFormat = "General"
                .Value = .Value
            End With
            
    End If
    
    If Err = 1004 Or Err = 94 Then
        MsgBox "A formula in your selection contains an error. Please correct this formula " & _
        "and restart the macro.", vbInformation
        Exit Sub
    ElseIf Err <> 0 Then
        MsgBox "Error " & Err & ": " & Error(Err.Number), vbCritical
        Exit Sub
    End If
    
    Selection.Replace What:="#VALUE!", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="#N/A", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False

Call UsageLog("Formulas to Values")

'MsgBox (WorkRange.Cells.Count & " formula(s) converted to values.")

End Sub

Sub DeleteSheet(control As IRibbonControl)
Dim ws As Worksheet
Application.DisplayAlerts = False

For Each ws In ActiveWindow.SelectedSheets

ws.Delete

Next ws


'Select Case ActiveSheet.Index
'Case Is = 1: ActiveSheet.Delete
'Case Else: Sheets(ActiveSheet.Index - 1).Activate
'    Sheets(ActiveSheet.Index + 1).Delete
'End Select

Application.DisplayAlerts = True

Call UsageLog("Delete Sheet")

End Sub

Sub SwapTrailingNegative(control As IRibbonControl)
Dim rng As Range
Dim WorkRng As Range
'Convert cells with a trailing negative to a negative value

Application.ScreenUpdating = False

On Error Resume Next
'Set Workrng
Select Case Selection.Count
    Case Is > 1: Set WorkRng = Selection.SpecialCells(xlCellTypeConstants)
    Case 1
        If IsNumeric(Selection) Then
            Set WorkRng = Selection
        End If
End Select

'Error Handling
    If Err <> 0 Then
        MsgBox "Please select a valid range.", vbCritical
        Exit Sub
    End If

'Perform Action
For Each rng In WorkRng

    If Right(rng, 1) = "-" And IsNumeric(rng) Then
        rng = -Left(rng, Len(rng) - 1)
    End If
    
    If IsNumeric(rng) Then
        rng.Style = "Comma"
    End If
    
Next

Call UsageLog("Swap Trailing Negative")

Application.ScreenUpdating = True

End Sub

Sub ShowDirectoryItems(control As IRibbonControl)
    Dim Directory As String
    Dim FileName As String
    Dim KeepExtensions As String
    Dim oFS As Object
    
    Set oFS = CreateObject("Scripting.FileSystemObject")

Application.ScreenUpdating = False

    Worksheets.Add

    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Select the directory path."
        .InitialFileName = CreateObject("WScript.Shell").SpecialFolders("Desktop")
        .Show
        
    Select Case .SelectedItems.Count
        Case Is = 0: GoTo EndEarly
        Case Is = 1: Directory = .SelectedItems(1) & "\"
    End Select
    
    End With

    KeepExtensions = MsgBox("Would you like to keep the file extensions?", vbYesNo, "Keep Extensions?")

    FileName = Dir(Directory) ' + vbHidden + vbSystem)
    
    Range("A1") = "File Name"
    Range("B1") = "File Size"
    Range("C1") = "Last Modified"
    
    Do While FileName <> ""
        ActiveCell.Offset(1, 0).Select
        ActiveCell.Offset(0, 1).Value = oFS.GetFile(Directory & FileName).Size
        ActiveCell.Offset(0, 2).Value = oFS.GetFile(Directory & FileName).DateLastModified
        If KeepExtensions = vbNo Then FileName = Left(FileName, InStrRev(FileName, ".") - 1)
        ActiveCell = FileName
        FileName = Dir
    Loop
    
If KeepExtensions = vbNo Then

    With Range(Selection, Selection.End(xlUp))
    .Replace What:=" ?-?-????", Replacement:="", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
    .Replace What:=" ??-?-????", Replacement:="", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
    .Replace What:=" ??-??-????", Replacement:="", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
    .Replace What:=" ?-??-????", Replacement:="", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
    End With
    
ActiveCell.End(xlUp).Select

End If
    
EndEarly:

    Set oFS = Nothing

Application.ScreenUpdating = True

End Sub


Sub RemoveLeadingSpaces()
Dim Cell As Range

For Each Cell In Selection
        Do Until Left(Cell, 1) <> " "
            Cell = Right(Cell, Len(Cell) - 1)
        Loop
Next Cell

End Sub

Sub MakeProper()
Dim Cell As Range

For Each Cell In Selection
    Cell.Value = WorksheetFunction.Proper(Cell)
Next
End Sub

Sub RoundTwo(control As IRibbonControl)
Dim Cell As Range

If Selection.Count = 1 Then
    If IsNumeric(Selection) And Not IsEmpty(Selection) Then
Selection = WorksheetFunction.Round(Selection, 2)
Exit Sub
    Else: Exit Sub
End If
End If

For Each Cell In Intersect(Selection.SpecialCells(xlCellTypeConstants), Selection.CurrentRegion)
    If IsNumeric(Cell.Value) Then
    Cell.Value = WorksheetFunction.Round(Cell, 2)
    Cell.NumberFormat = "0.00"
    End If

Next
End Sub


Sub GoToRow(control As IRibbonControl, GoToRow As Integer)

'RowNumber = InputBox("What row number would you like to go to?", "RowNum")
Cells(GoToRow, 1).EntireRow.Select

Do Until Intersect(ActiveWindow.VisibleRange, ActiveCell) Is Nothing
ActiveWindow.SmallScroll Down:=1
Loop

ActiveWindow.SmallScroll Up:=1

End Sub

Sub PasteAndRemoveDuplicates(control As IRibbonControl)

On Error Resume Next

    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Selection.RemoveDuplicates Columns:=1, Header:=xlNo

End Sub

Sub RemoveTrailingSpaces(control As IRibbonControl)
Dim Cell As Range

For Each Cell In Intersect(Selection.SpecialCells(xlCellTypeConstants), Selection.Parent.UsedRange)

    Do Until Right(Cell, 1) <> " "
        Cell = Left(Cell, Len(Cell) - 1)
    Loop

Next

End Sub

Sub CountDuplicates(control As IRibbonControl)
Dim Cell As Range
Dim WorkRange As Range
Dim UniqueCount As String
Dim UniqueItems As Integer

    If Selection.Cells.Count = 1 Then
    
    Exit Sub
    
    ElseIf Selection.Cells.Count > 1 Then

        Set WorkRange = Selection.SpecialCells(xlConstants)
        Set WorkRange = Intersect(WorkRange, WorkRange.Parent.UsedRange)
    
    End If
    
    If WorkRange Is Nothing Then: MsgBox "Please select a valid range.": Exit Sub

For Each Cell In WorkRange
    If Cell.Row <> 1 Then
    If InStr(1, UniqueCount, Cell) = 0 Then
    UniqueCount = UniqueCount & Cell.Value & " " & Application.WorksheetFunction.CountIf(WorkRange, Cell) & vbNewLine
    UniqueItems = UniqueItems + 1
    End If
    End If
Next Cell

MsgBox ("Number of Unique Items: " & UniqueItems)

End Sub

Sub CountListDuplicates() 'control As IRibbonControl
Dim Cell As Range
Dim WorkRange As Range
Dim UniqueCount As String
Dim UniqueItems As Integer

    If Selection.Cells.Count = 1 Then
    
    Exit Sub
    
    ElseIf Selection.Cells.Count > 1 Then

        Set WorkRange = Selection.SpecialCells(xlConstants)
        Set WorkRange = Intersect(WorkRange, WorkRange.Parent.UsedRange)
    
    End If
    
    If WorkRange Is Nothing Then: MsgBox "Please select a valid range.": Exit Sub

For Each Cell In WorkRange
    If Cell.Row <> 1 Then
    If InStr(1, UniqueCount, Cell) = 0 Then
    UniqueCount = UniqueCount & Cell.Value & " " & Application.WorksheetFunction.CountIf(WorkRange, Cell) & vbNewLine
    UniqueItems = UniqueItems + 1
    End If
    If Len(UniqueCount) > 1000 Then: MsgBox ("This selection has way too many unique values."): Exit Sub
    End If
Next Cell

MsgBox (UniqueCount & vbNewLine & "Number of Unique Items: " & UniqueItems)

End Sub


Sub ListCriteriaTest()

Dim Cell As Range
Dim Message As String
Dim List As String
Dim clipboard As MSForms.DataObject
Set clipboard = New MSForms.DataObject

On Error Resume Next

    If Selection.Cells.Count = 1 Then
        If Not IsEmpty(ActiveCell) And Not ActiveCell.HasFormula Then
        
        Do Until Len(Message) - 3 > 245
            If ActiveCell.EntireRow.Hidden <> True Then
            If InStr(1, ActiveCell, Cell) = 0 And Not IsEmpty(ActiveCell) Then
                Message = Message & ActiveCell.Value & "','"
            End If
            End If
        ActiveCell.Offset(1, 0).Select
        
        If IsEmpty(ActiveCell) Then Exit Do
        
        Loop
        
        End If
        
    End If
            
If Len(Left(Message, Len(Message) - 3)) > 255 Then
MsgBox ("Each list member in PeopleSoft can contain a maximum of 28 EmplIDs." & vbNewLine & vbNewLine & _
    "Please shrink your selection size and then add multiple selections to the PeopleSoft list members."), vbCritical
Exit Sub
End If

Message = Left(Message, Len(Message) - 3)

clipboard.SetText Message
clipboard.PutInClipboard

MsgBox ("This list has been copied to your clipboard." & vbNewLine & vbNewLine & _
        "You can paste it into the PeopleSoft list members to add.")
        
Call UsageLog("List Criteria")

End Sub

Sub SectionData(control As IRibbonControl)
Dim Continue As String

If ActiveCell.End(xlDown).Row > 1000 Then
Continue = MsgBox("Over 1000 records found." & vbNewLine & vbNewLine & "Do you want to proceed?", vbYesNo)
If Continue = vbNo Then Exit Sub
End If

Application.ScreenUpdating = False

Do Until IsEmpty(ActiveCell)

    If ActiveCell.Row = 1 Then ActiveCell.Offset(1, 0).Select
    If ActiveCell.Row = 2 Then ActiveCell.Offset(1, 0).Select
        
    If ActiveCell = ActiveCell.Offset(-1, 0) Or IsEmpty(ActiveCell.Offset(-1, 0)) Then
    
    ElseIf ActiveCell <> ActiveCell.Offset(-1, 0) Then
        Rows(ActiveCell.Row & ":" & ActiveCell.Row).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    End If
    
        ActiveCell.Offset(1, 0).Select
Loop

Application.ScreenUpdating = True

Cells(1, ActiveCell.Column).Select

Call UsageLog("Section Data")

End Sub

Sub test()

Dim BUSheet As Integer

For BUSheet = 1 To ActiveWorkbook.Sheets.Count
Sheets(BUSheet).Name = Replace(Sheets(BUSheet).Name, "Business Unit = ", "")
Next BUSheet

End Sub

Sub FormatCharts(control As IRibbonControl)
Dim ChartNum As Integer

If ActiveSheet.ChartObjects.Count < 1 Then Exit Sub
ActiveSheet.ChartObjects(1).Activate

For ChartNum = 1 To ActiveSheet.ChartObjects.Count

ActiveSheet.ChartObjects(ChartNum).Height = ActiveChart.Parent.Height
ActiveSheet.ChartObjects(ChartNum).Width = ActiveChart.Parent.Width

ActiveSheet.ChartObjects(ChartNum).Top = 30

Next ChartNum

Application.SendKeys "{Esc}"

End Sub

Sub ConvertPhoneNumbers(control As IRibbonControl)

With Intersect(Selection, ActiveSheet.UsedRange)
    .Replace What:="(", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    .Replace What:=")", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    .Replace What:="-", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    .Replace What:=" ", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    .NumberFormat = "[<=9999999]###-####;(###) ###-####"
End With

End Sub

Sub GoToBlanks(control As IRibbonControl)
Selection.SpecialCells(xlCellTypeBlanks).Select
End Sub

Sub RemoveDates(control As IRibbonControl)

With Intersect(Selection, ActiveSheet.UsedRange)
    .Replace What:=" ??-??-????", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    .Replace What:=" ?-??-????", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    .Replace What:=" ??-?-????", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    .Replace What:=" ?-?-????", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    .Replace What:=".txt", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
End With

End Sub

Sub SystemDirectoryPaths()
'Print the user's default AddIn folder
Debug.Print "C:\Users\" & Environ$("Username") & "\AppData\Roaming\Microsoft\AddIns"

'Print the user's desktop path
Debug.Print CreateObject("WScript.Shell").SpecialFolders("Desktop")
End Sub


Sub RemoveDatesAndExtendtions(control As IRibbonControl)

    Selection.Replace What:=" ??-??-????", Replacement:="", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:=" ?-??-????", Replacement:="", LookAt:=xlPart _
        , SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:=" ??-?-????", Replacement:="", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:=" ?-?-????", Replacement:="", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:=".txt", Replacement:="", LookAt:=xlPart _
        , SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:=".pdf", Replacement:="", LookAt:=xlPart _
        , SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:=".xls", Replacement:="", LookAt:=xlPart _
        , SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:=".xlsx", Replacement:="", LookAt:=xlPart _
        , SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
End Sub

Sub ListProcedures()
    Dim VBP As VBIDE.VBProject
    Dim VBC As VBComponent
    Dim CM As CodeModule
    Dim StartLine As Long
    Dim Msg As String
    Dim ProcName As String
    Dim newWB As Workbook
    Dim rowIndex As Integer
    Set newWB = Workbooks.Add
    
'   Use the workbook this code is in
    Set VBP = ThisWorkbook.VBProject
    
'   Loop through the VB components
    For Each VBC In VBP.VBComponents
        Set CM = VBC.CodeModule
        Msg = Msg & vbNewLine
        StartLine = CM.CountOfDeclarationLines + 1
        Do Until StartLine >= CM.CountOfLines
        rowIndex = rowIndex + 1
            newWB.Sheets(1).Cells(rowIndex, 1).Value = VBC.Name
            newWB.Sheets(1).Cells(rowIndex, 2).Value = CM.ProcOfLine(StartLine, vbext_pk_Proc)
            StartLine = StartLine + CM.ProcCountLines _
              (CM.ProcOfLine(StartLine, vbext_pk_Proc), vbext_pk_Proc)
        Loop
    Next VBC
    
    newWB.Sheets(1).Columns.AutoFit

End Sub

