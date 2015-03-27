Attribute VB_Name = "PivotTables"
Option Explicit

Sub ConvertPivotCall()

    ActiveSheet.PivotTables(ActiveCell.PivotTable.Name).PivotSelect "", xlDataAndLabel, True
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Selection.Copy
    Selection.Cells(1).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Selection.Cells(1).Select

End Sub

Sub NormalizePivotData()
If ActiveSheet.PivotTables.Count <> 1 Then Exit Sub
If Left(ActiveSheet.PivotTables(1).TableRange1.Cells(1).Address, 4) <> "$A$3" Then Exit Sub

Dim pt As PivotTable
Dim pf As PivotField
Dim ColumnCount As Integer
Dim cell As Range
Dim counter As Integer
Dim FillCells As Integer


Set pt = ActiveSheet.PivotTables(1)

    pt.ColumnGrand = False
    pt.RowGrand = False

For Each pf In pt.PivotFields
    If pf.Orientation = xlRowField Then
    Range("A1").Insert Shift:=xlToRight
    Range("A1") = pf.Caption
    ColumnCount = ColumnCount + 1
    End If

Next pf

ActiveSheet.PivotTables(1).TableRange1.Cells(1).Select

Call ConvertPivotCall

For Each cell In Range(ActiveCell.Address, ActiveCell.End(xlDown))

If cell.IndentLevel > 0 Then

    Do While cell.IndentLevel - counter > 0
        cell.Insert Shift:=xlToRight
    counter = counter + 1
    Loop
    counter = 0
    
Else

Range(cell.Offset(0, 1), cell.Offset(0, 10)).ClearContents

End If

Next cell

For Each cell In Intersect(Columns(ColumnCount), ActiveSheet.UsedRange)
If IsEmpty(cell.Offset(0, -1)) Then

FillCells = 1
Do Until FillCells = ColumnCount

Cells(cell.Row, FillCells) = Cells(cell.Row, FillCells).End(xlUp)
FillCells = FillCells + 1

Loop

Else

'Rows(Cell.Row).Delete

End If
Next cell


End Sub

Sub test()
Dim pf As PivotField
Dim pt As PivotTable
Dim ColumnCount As Integer
Dim cell As Range
Dim CellValue As Range
Dim IndentCount As Integer

Set pt = ActiveSheet.PivotTables(1)

    pt.ColumnGrand = False
    pt.RowGrand = False

For Each pf In pt.PivotFields
    If pf.Orientation = xlRowField Then
    Cells(3, 2 + pf.Position) = pf.Caption
    ColumnCount = ColumnCount + 1
    End If

Next pf

Cells(3, 3 + ColumnCount) = "Count"

For Each cell In Range(Range("C3").Offset(1, ColumnCount), ActiveSheet.UsedRange.SpecialCells(xlCellTypeLastCell).Address)

'If the cell in column 1 on the corresponding row is a data point
If Cells(cell.Row, 1).IndentLevel = ColumnCount - 1 Then

For IndentCount = 1 To ColumnCount

Set CellValue = Cells(cell.Row, 1)
Do Until CellValue.IndentLevel = IndentCount - 1
Set CellValue = CellValue.Offset(-1, 0)
Loop
cell.Offset(0, -ColumnCount + IndentCount - 1) = CellValue

Next IndentCount

cell = cell.Offset(0, -ColumnCount - 1)

Else

End If

Next cell

Range("C4", ActiveSheet.UsedRange.SpecialCells(xlCellTypeLastCell).Address).SpecialCells(xlCellTypeBlanks).Delete Shift:=xlUp

Columns.AutoFit

End Sub

Sub test1()

Dim pt As PivotTable
Dim pf As PivotField

Set pt = ActiveSheet.PivotTables(1)

    pt.ColumnGrand = False
    pt.RowGrand = False

Do Until 0 = 1

For Each pf In pt.PivotFields
Debug.Print pf.Orientation
Debug.Print pf.Caption
Debug.Print pf.Position

Next pf

Loop

End Sub

