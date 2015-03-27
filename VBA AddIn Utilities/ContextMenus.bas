Attribute VB_Name = "ContextMenus"
Option Explicit

Sub ConvertPivot(control As IRibbonControl)

    Dim pt As PivotTable
    Dim col As PivotFields
    Dim c As PivotField
    
Application.ScreenUpdating = False
    
    With ActiveSheet.PivotTables(ActiveCell.PivotTable.Name)
        .RowAxisLayout xlTabularRow
        .ColumnGrand = False
        .RowGrand = False
    End With

    Set pt = ActiveSheet.PivotTables(ActiveCell.PivotTable.Name)
    Set col = pt.PivotFields
On Error Resume Next
    For Each c In col
        pt.PivotFields(c.Name).Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    Next
On Error GoTo 0

    ActiveCell.CurrentRegion.Copy
    Sheets.Add After:=ActiveSheet
    
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    Do Until Range("A1") <> ""
    Rows(1).Delete
    Loop
    
    Intersect(ActiveCell.CurrentRegion, Columns(1)).Select

On Error GoTo skip

    Selection.SpecialCells(xlCellTypeBlanks).Select
    Selection.FormulaR1C1 = "=R[-1]C"
    
skip:
On Error GoTo 0
    
    Columns(1).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("A1").Select
    If Left(ActiveCell.Value, 3) = "Sum" Then Rows(1).Delete
    
    ActiveCell.CurrentRegion.Replace What:="(blank)", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False

    Intersect(Rows(1), ActiveSheet.UsedRange).EntireColumn.AutoFit
    Application.CutCopyMode = False
    
Sheets(ActiveSheet.Index - 1).Select

    With ActiveSheet.PivotTables(ActiveCell.PivotTable.Name)
        .RowAxisLayout xlCompactRow
        .ColumnGrand = False
        .RowGrand = False
    End With
    
Sheets(ActiveSheet.Index + 1).Select
    
Application.ScreenUpdating = True


End Sub

Sub PasteAndDeliminateComma(control As IRibbonControl)

Application.ScreenUpdating = False

    ActiveSheet.Paste
        
On Error Resume Next

    Selection.TextToColumns _
      Destination:=ActiveCell, _
      DataType:=xlDelimited, _
      TextQualifier:=xlDoubleQuote, _
      ConsecutiveDelimiter:=False, _
      Tab:=False, _
      Semicolon:=False, _
      Comma:=True, _
      Space:=False, _
      OtherChar:=""
      
      
Selection.TextToColumns Destination:=ActiveCell, DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=True, Space:=False, Other:=True, OtherChar:= _
        "", FieldInfo:=Array(Array(1, 1), Array(2, 1)), TrailingMinusNumbers:=True

Application.ScreenUpdating = True

End Sub

Sub PasteAndDeliminateSpace(control As IRibbonControl)

Application.ScreenUpdating = False

    ActiveSheet.Paste
        
On Error Resume Next
        
    Selection.TextToColumns _
      Destination:=ActiveCell, _
      DataType:=xlDelimited, _
      TextQualifier:=xlDoubleQuote, _
      ConsecutiveDelimiter:=False, _
      Tab:=True, _
      Semicolon:=False, _
      Comma:=False, _
      Space:=True, _
      Other:=False

Application.ScreenUpdating = True

End Sub

Sub TexttoFormula(control As IRibbonControl)

If Selection.Count <> 1 Then
MsgBox "This macro can only be run for a single cell."
Exit Sub
End If

On Error Resume Next

ActiveCell.NumberFormat = "General"
ActiveCell.Formula = ActiveCell.Formula
Application.SendKeys "{F2}"
Application.SendKeys "{Enter}"

Call UsageLog("Text to Formula")

End Sub


Sub Trimed(control As IRibbonControl)
Dim cell As Range

If Selection.Cells.Count <= 1 Then
Selection = Application.WorksheetFunction.Trim(Selection)
Else
For Each cell In Selection.SpecialCells(xlCellTypeConstants)
cell = Application.WorksheetFunction.Trim(cell)
Next cell
End If

End Sub


Sub CopySum(control As IRibbonControl)
Dim SelectionSum As Double
Dim clipboard As MSForms.DataObject
Set clipboard = New MSForms.DataObject

SelectionSum = Application.WorksheetFunction.Sum(Selection)

clipboard.SetText SelectionSum
clipboard.PutInClipboard

End Sub

Sub ProperCase(control As IRibbonControl)
Dim cell As Range

For Each cell In Intersect(ActiveSheet.UsedRange, Selection)
cell = Application.WorksheetFunction.Proper(cell)
Next cell

End Sub
