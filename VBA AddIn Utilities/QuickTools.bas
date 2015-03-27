Attribute VB_Name = "QuickTools"
Option Explicit
Dim workRange As Range

Sub ConvertEmpIDToText(control As IRibbonControl)
Dim cell As Range
Dim EmpIDCount As Integer

Application.ScreenUpdating = False

On Error Resume Next

    If Selection.Cells.Count = 1 Then
        If Not IsEmpty(ActiveCell) And Not ActiveCell.HasFormula Then
            
            Set workRange = Selection
        
        Else

            MsgBox "Please select a range with constants", vbInformation
            Application.ScreenUpdating = True
            Exit Sub
        
        End If
    
    ElseIf Selection.Cells.Count > 1 Then

        Set workRange = Selection.SpecialCells(xlCellTypeConstants)
        Set workRange = Intersect(workRange, workRange.Parent.UsedRange)
    
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
    
    If workRange Is Nothing Then
        MsgBox "Please select a valid range."
        Application.ScreenUpdating = True
        Exit Sub
    End If

For Each cell In workRange

    Do Until Right(cell, 1) <> " "
        cell = Left(cell, Len(cell) - 1)
    Loop

If IsNumeric(cell) Then

    cell.NumberFormat = "@"
    
    If Len(cell) = 1 Then
        cell.Value = "00000" & cell.Value
    ElseIf Len(cell) = 2 Then
        cell.Value = "0000" & cell.Value
    ElseIf Len(cell) = 3 Then
        cell.Value = "000" & cell.Value
    ElseIf Len(cell) = 4 Then
        cell.Value = "00" & cell.Value
    ElseIf Len(cell) = 5 Then
        cell.Value = "0" & cell.Value
    End If
    
    If Len(cell) = 6 Then
        cell.Value = "'" & cell.Value
    End If

End If
        
Next cell

Call UsageLog("EmpID to Text")

Application.ScreenUpdating = True

'MsgBox (EmpIDCount & " EmpID(s) converted to text.")

End Sub

Sub ConvertSSNToText(control As IRibbonControl)
Dim cell As Range
Dim workRange As Range
Dim SSNCount As Integer

On Error Resume Next

    If Selection.Cells.Count = 1 Then
        If Not IsEmpty(ActiveCell) And Not ActiveCell.HasFormula Then
            
            Set workRange = Selection
        
        Else

            MsgBox "Please select a range with constants", vbInformation
            Exit Sub
        
        End If
    
    ElseIf Selection.Cells.Count > 1 Then

        Set workRange = Selection.SpecialCells(xlCellTypeConstants)
        Set workRange = Intersect(workRange, workRange.Parent.UsedRange)
    
    End If
    
    If Err = 1004 Or Err = 94 Or Err = 91 Then
        MsgBox "There are no constants in your selection", vbInformation
        Exit Sub
    ElseIf Err <> 0 Then
        MsgBox "Error " & Err & ": " & Error(Err.Number)
        Exit Sub
    End If
    
On Error GoTo 0
    
    If workRange Is Nothing Then
        MsgBox "Please select a valid range."
        Exit Sub
    End If

workRange.Replace What:="-", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False

For Each cell In workRange

If IsNumeric(cell) Then

    cell.NumberFormat = "@"
    
        If Len(cell) = 6 Then
            cell.Value = "000" & cell.Value
        ElseIf Len(cell) = 7 Then
            cell.Value = "00" & cell.Value
        ElseIf Len(cell) = 8 Then
            cell.Value = "0" & cell.Value
        End If

        If Len(cell) = 9 Then
            cell.Value = "'" & cell.Value
            SSNCount = SSNCount + 1
        End If
End If
        
Next cell

Call UsageLog("SSN to Text")

'MsgBox (SSNCount & " SSN(s) converted to text.")

End Sub

Sub PSTable(control As IRibbonControl)
Dim TableName As String
Dim ColumnCount As Integer
Dim ActiveColumn As Integer
Dim cell As Range
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

For Each cell In Range("A1:" & Cells(1, Range("A1").End(xlToRight).Column).Address)
        
    cell.HorizontalAlignment = cell.Offset(1, 0).HorizontalAlignment
    
    If IsNumeric(cell.Offset(1, 0)) Or IsDate(cell.Offset(1, 0)) Then cell.HorizontalAlignment = xlRight
    If cell.Offset(1, 0).NumberFormat = "@" Then cell.HorizontalAlignment = xlLeft

cell.VerticalAlignment = xlCenter

Next cell

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

Dim workRange As Range
Dim cell As Range
Dim List As String

Dim clipboard As MSForms.DataObject
Set clipboard = New MSForms.DataObject

If TypeName(Selection) <> "Range" Then Exit Sub

Application.ScreenUpdating = False

If Selection.Columns.Count > 1 Then
    MsgBox "Please select a single column."
    Exit Sub
End If

Set workRange = Intersect(Range(ActiveCell, ActiveCell.End(xlDown)).SpecialCells(xlCellTypeVisible), ActiveSheet.UsedRange)

For Each cell In workRange

    If InStr(1, List, cell) = 0 And Not IsEmpty(cell) Then
    
        cell = Replace(cell, " ", "")
        
        If Len(cell) < 6 Then cell = numberToEmpID(cell)
    
        If Len(List & cell.Value) <= 255 Then
        
            List = List & cell.Value & "','"
        
        Else
        
            Exit For
            
        End If
    
    End If
    
    cell.Select

Next cell

'Truncate the ',' from the end of the list
List = Left(List, Len(List) - 3)

'If we've reached the end of the used range, send the active cell to the top of the worksheet
If Intersect(ActiveCell.Offset(1, 0), ActiveSheet.UsedRange) Is Nothing Then
    ActiveCell.End(xlUp).Select
End If

clipboard.SetText List
clipboard.PutInClipboard

Application.ScreenUpdating = True

MsgBox ("This list has been copied to your clipboard." & vbNewLine & vbNewLine & _
        "Paste this list in the 'List Members' of a PeopleSoft Query criteria for bulk processing.")

End Sub

Private Function numberToEmpID(cell As Range) As Range

If IsNumeric(cell) Then

    cell.NumberFormat = "@"
    
    If Len(cell) = 1 Then
        cell.Value = "00000" & cell.Value
    ElseIf Len(cell) = 2 Then
        cell.Value = "0000" & cell.Value
    ElseIf Len(cell) = 3 Then
        cell.Value = "000" & cell.Value
    ElseIf Len(cell) = 4 Then
        cell.Value = "00" & cell.Value
    ElseIf Len(cell) = 5 Then
        cell.Value = "0" & cell.Value
    End If
    
    cell.Value = "'" & cell.Value

End If

Set numberToEmpID = cell

End Function


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
Dim cell As Range
Dim workRange As Range
Dim NameCounter As Integer

On Error Resume Next

    If Selection.Cells.Count = 1 Then
        If Not IsEmpty(ActiveCell) And Not ActiveCell.HasFormula Then
            
            Set workRange = Selection
        
        Else

            MsgBox "Please select a range with names in 'Last, First' format.", vbInformation
            Exit Sub
        
        End If
    
    ElseIf Selection.Cells.Count > 1 Then

        Set workRange = Selection.SpecialCells(xlCellTypeConstants)
        Set workRange = Intersect(workRange, workRange.Parent.UsedRange)
    
    End If
    
    If Err = 1004 Or Err = 94 Or Err = 91 Then
        MsgBox "There are no constants in your selection", vbInformation
        Exit Sub
    
    ElseIf Err = 6 Then
    
        Set workRange = Selection.SpecialCells(xlCellTypeConstants)
        Set workRange = Intersect(workRange, workRange.Parent.UsedRange)
    
    ElseIf Err <> 0 Then
        MsgBox "Error " & Err & ": " & Error(Err.Number), vbCritical
        Exit Sub
    End If

ErrorHandler:

On Error GoTo 0


Dim NameValue As Boolean
For Each cell In workRange

    
    NameValue = cell Like "[$,;,:]"
    If NameValue <> True Then
        
        If InStr(cell, ",") Then
            
            LastName = Left(cell, InStr(cell, ",") - 1)
            FirstName = Right(cell, Len(cell) - InStr(cell, ","))
        
            If Left(FirstName, 1) = " " Then
                FirstName = Right(FirstName, Len(FirstName) - 1)
            End If
            
            cell = WorksheetFunction.Proper(FirstName & " " & LastName)

            cell = Left(cell, InStr(cell, " ") - 1) & " " & Right(cell, Len(cell) - InStrRev(cell, " "))
            
            NameCounter = NameCounter + 1
        
        End If
        
        If InStr(cell, "  ") Then
            
            LastName = Left(cell, InStr(cell, "  ") - 1)
            FirstName = Right(cell, Len(cell) - InStr(cell, "  "))
        
            If Left(FirstName, 1) = " " Then
                FirstName = Right(FirstName, Len(FirstName) - 1)
            End If
            
            cell = WorksheetFunction.Proper(FirstName & " " & LastName)
            
            cell = Left(cell, InStr(cell, " ") - 1) & " " & Right(cell, Len(cell) - InStrRev(cell, " "))
            
            NameCounter = NameCounter + 1
        
        End If
        
        Do Until Left(cell, 1) <> " "
            cell = Right(cell, Len(cell) - 1)
        Loop
    
    End If

Next cell
Call UsageLog("Reverse Name")

'MsgBox NameCounter & " name(s) were reformatted to 'First Last' format."

End Sub

Sub CopyWorkSheet(control As IRibbonControl)

If IsEmpty(ActiveSheet) Then Exit Sub

ActiveSheet.Copy After:=ActiveSheet

Call UsageLog("Copy Worksheet")

End Sub

Sub FormulastoValues(control As IRibbonControl)
Dim workRange As Range
Dim intArea As Integer
Dim rngCell As Range
'Dim continue As String
'
'continue = MsgBox("Would you like to convert all formulas in your selection to values?", _
'            vbYesNo, "Convert Formulas to Values")
'
'    If continue = vbNo Then Exit Sub
    
    Set workRange = Selection
    Set workRange = Intersect(workRange, workRange.Parent.UsedRange)

On Error Resume Next

    If Not workRange.Cells.Count = 1 Then
        Set workRange = workRange.SpecialCells(xlCellTypeFormulas)
    End If
    
    If Err = 1004 Or Err = 94 Then
        MsgBox "There are no formulas in your selection", vbInformation
        Exit Sub
    ElseIf Err <> 0 Then
        MsgBox "Error " & Err & ": " & Error(Err.Number), vbCritical
        Exit Sub
    End If
    
On Error GoTo 0
    
    If workRange Is Nothing Then
        MsgBox "Please select a valid range.", vbCritical
        Exit Sub
    End If

    If workRange.Cells.Count = 1 And IsEmpty(Selection.Cells(1, 1)) Then
        MsgBox "Please select a range with values", vbInformation
        Exit Sub
    End If

On Error Resume Next

    If workRange.Areas.Count > 1 Then
    
        For intArea = 1 To workRange.Areas.Count
            With workRange.Areas(intArea)
                'Debug.Print WorkRange.Areas(intArea).Address
               ' .NumberFormat = "General"
                .Value = .Value
            End With
        Next
        
    ElseIf workRange.Areas.Count = 1 Then
    
            With workRange
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
Dim cell As Range

For Each cell In Selection
        Do Until Left(cell, 1) <> " "
            cell = Right(cell, Len(cell) - 1)
        Loop
Next cell

End Sub

Sub MakeProper()
Dim cell As Range

For Each cell In Selection
    cell.Value = WorksheetFunction.Proper(cell)
Next
End Sub

Sub RoundTwo(control As IRibbonControl)
Dim cell As Range

If Selection.Count = 1 Then
    If IsNumeric(Selection) And Not IsEmpty(Selection) Then
Selection = WorksheetFunction.Round(Selection, 2)
Exit Sub
    Else: Exit Sub
End If
End If

For Each cell In Intersect(Selection.SpecialCells(xlCellTypeConstants), Selection.CurrentRegion)
    If IsNumeric(cell.Value) Then
    cell.Value = WorksheetFunction.Round(cell, 2)
    cell.NumberFormat = "0.00"
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
Dim cell As Range

For Each cell In Intersect(Selection.SpecialCells(xlCellTypeConstants), Selection.Parent.UsedRange)

    Do Until Right(cell, 1) <> " "
        cell = Left(cell, Len(cell) - 1)
    Loop

Next

End Sub

Sub CountDuplicates(control As IRibbonControl)
Dim cell As Range
Dim workRange As Range
Dim UniqueCount As String
Dim UniqueItems As Integer

    If Selection.Cells.Count = 1 Then
    
    Exit Sub
    
    ElseIf Selection.Cells.Count > 1 Then

        Set workRange = Selection.SpecialCells(xlConstants)
        Set workRange = Intersect(workRange, workRange.Parent.UsedRange)
    
    End If
    
    If workRange Is Nothing Then: MsgBox "Please select a valid range.": Exit Sub

For Each cell In workRange
    If cell.Row <> 1 Then
    If InStr(1, UniqueCount, cell) = 0 Then
    UniqueCount = UniqueCount & cell.Value & " " & Application.WorksheetFunction.CountIf(workRange, cell) & vbNewLine
    UniqueItems = UniqueItems + 1
    End If
    End If
Next cell

MsgBox ("Number of Unique Items: " & UniqueItems)

End Sub

Sub CountListDuplicates() 'control As IRibbonControl
Dim cell As Range
Dim workRange As Range
Dim UniqueCount As String
Dim UniqueItems As Integer

    If Selection.Cells.Count = 1 Then
    
    Exit Sub
    
    ElseIf Selection.Cells.Count > 1 Then

        Set workRange = Selection.SpecialCells(xlConstants)
        Set workRange = Intersect(workRange, workRange.Parent.UsedRange)
    
    End If
    
    If workRange Is Nothing Then: MsgBox "Please select a valid range.": Exit Sub

For Each cell In workRange
    If cell.Row <> 1 Then
    If InStr(1, UniqueCount, cell) = 0 Then
    UniqueCount = UniqueCount & cell.Value & " " & Application.WorksheetFunction.CountIf(workRange, cell) & vbNewLine
    UniqueItems = UniqueItems + 1
    End If
    If Len(UniqueCount) > 1000 Then: MsgBox ("This selection has way too many unique values."): Exit Sub
    End If
Next cell

MsgBox (UniqueCount & vbNewLine & "Number of Unique Items: " & UniqueItems)

End Sub


Sub ListCriteriaTest()

Dim cell As Range
Dim Message As String
Dim List As String
Dim clipboard As MSForms.DataObject
Set clipboard = New MSForms.DataObject

On Error Resume Next

    If Selection.Cells.Count = 1 Then
        If Not IsEmpty(ActiveCell) And Not ActiveCell.HasFormula Then
        
        Do Until Len(Message) - 3 > 245
            If ActiveCell.EntireRow.Hidden <> True Then
            If InStr(1, ActiveCell, cell) = 0 And Not IsEmpty(ActiveCell) Then
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

