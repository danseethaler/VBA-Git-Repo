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

Sub ColumnsAutofit(control As IRibbonControl)
Dim MinWidth As Integer
Dim MaxWidth As Integer
Dim ColumnCount As Integer
Dim ActiveColumn As Integer

ColumnCount = Range("A1:" & ActiveSheet.UsedRange.SpecialCells(xlCellTypeLastCell).Address).Columns.Count

On Error Resume Next

Columns.AutoFit

If Err = 1004 Then
MsgBox "The autofit function will not work on this sheet. It may be protected.", vbInformation
Exit Sub
ElseIf Err <> 0 Then
MsgBox "Error: " & Err.Number
End If

'Set Variables
    MinWidth = 12
    MaxWidth = 35

For ActiveColumn = 1 To ColumnCount Step 1

    If Columns(ActiveColumn).ColumnWidth < MinWidth Then
        'Columns(ActiveColumn).ColumnWidth = MinWidth
    ElseIf Columns(ActiveColumn).ColumnWidth > MaxWidth Then
        Columns(ActiveColumn).ColumnWidth = MaxWidth
    End If

If Columns(ActiveColumn).ColumnWidth < 35 Then _
Columns(ActiveColumn).ColumnWidth = Columns(ActiveColumn).ColumnWidth + 2

Next

Rows.AutoFit

Call UsageLog("Autofit Columns")

End Sub

Sub DateTimeFormat(control As IRibbonControl)

    Selection.NumberFormat = "[$-409]m/d/yy h:mm AM/PM;@"
    
    Call UsageLog("DateTimeFormat")

End Sub


Sub RemoveHyperlinks(control As IRibbonControl)
Dim Cell As Range
    For Each Cell In ActiveCell.CurrentRegion
    Cell.Hyperlinks.Delete
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
