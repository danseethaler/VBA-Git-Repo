Attribute VB_Name = "Settings"
Option Explicit

Sub RCToggle(control As IRibbonControl)

    If Application.ReferenceStyle = xlA1 Then
        Application.ReferenceStyle = xlR1C1
    Else
        Application.ReferenceStyle = xlA1
    End If
    
Call UsageLog("RC Toggle")
    
End Sub

Sub HideError(control As IRibbonControl)
'This sub will toggle the cell errors
Application.ErrorCheckingOptions.BackgroundChecking = Not _
Application.ErrorCheckingOptions.BackgroundChecking

Application.ScreenUpdating = True
    
Call UsageLog("Hide Error")

End Sub

Sub DateTimeFormat(control As IRibbonControl)

    Selection.NumberFormat = "[$-409]m/d/yy h:mm AM/PM;@"
    
    Call UsageLog("DateTimeFormat")

End Sub


Sub RemoveHyperlinks(control As IRibbonControl)
Dim cell As Range
    For Each cell In ActiveCell.CurrentRegion
    cell.Hyperlinks.Delete
    Next
    
Call UsageLog("Remove Hyperlinks")

End Sub

Sub IdentifyDuplicates()
    Selection.FormatConditions.AddUniqueValues
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    Selection.FormatConditions(1).DupeUnique = xlDuplicate
    With Selection.FormatConditions(1).Font
        .Color = -16383844
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13551615
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    
Call UsageLog("Identify Duplicates")

End Sub

Sub FreezeRow1()
Dim SelectedRange As Range

Set SelectedRange = Selection

If ActiveWindow.FreezePanes = True Then

    ActiveWindow.FreezePanes = False
    ActiveWindow.Split = False
    
Else
    Range("A1").Select
    
    With ActiveWindow
        .SplitColumn = 0
        .SplitRow = 1
    End With
    ActiveWindow.FreezePanes = True

End If

SelectedRange.Select

Call UsageLog("Freeze Row 1")

End Sub

Sub ShowDuplicates(control As IRibbonControl)
    Selection.FormatConditions.AddUniqueValues
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    Selection.FormatConditions(1).DupeUnique = xlDuplicate
    With Selection.FormatConditions(1).Font
        .Color = -16383844
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13551615
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False

Call UsageLog("Show Duplicates")

End Sub

Sub RoundCommaStyle(control As IRibbonControl)

    Intersect(Selection, Range("A2:" & ActiveSheet.UsedRange. _
    SpecialCells(xlCellTypeLastCell).Address)).NumberFormat = "_(* #,##0_);_(* (#,##0);_(* ""-""??_);_(@_)"

End Sub

Sub KillWorkbook(control As IRibbonControl)
'Save the current workbook and close it.

    Application.DisplayAlerts = False
        ActiveWorkbook.Close (True)
    Application.DisplayAlerts = True
    
If Workbooks.Count = 0 Then Application.Quit

End Sub


Sub DeleteSheet(control As IRibbonControl)
Dim ws As Worksheet
Application.DisplayAlerts = False

For Each ws In ActiveWindow.SelectedSheets

ws.Delete

Next ws

Application.DisplayAlerts = True

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

Application.ScreenUpdating = True

End Sub

Sub ShowDirectoryItems(control As IRibbonControl)
    Dim Directory As String
    Dim fileName As String
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

    fileName = Dir(Directory) ' + vbHidden + vbSystem)
    
    Range("A1") = "File Name"
    Range("B1") = "File Size"
    Range("C1") = "Last Modified"
    
    Do While fileName <> ""
        ActiveCell.Offset(1, 0).Select
        ActiveCell.Offset(0, 1).Value = oFS.GetFile(Directory & fileName).Size
        ActiveCell.Offset(0, 2).Value = oFS.GetFile(Directory & fileName).DateLastModified
        If KeepExtensions = vbNo Then fileName = Left(fileName, InStrRev(fileName, ".") - 1)
        ActiveCell = fileName
        fileName = Dir
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

Sub MakeProper()
Dim cell As Range

For Each cell In Selection
    cell.Value = WorksheetFunction.Proper(cell)
Next
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

Err.Clear
On Error Resume Next

Selection.SpecialCells(xlCellTypeBlanks).Select

If Err.Number = "1004" Then MsgBox "No blank cells were found.", Buttons:=vbInformation

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
