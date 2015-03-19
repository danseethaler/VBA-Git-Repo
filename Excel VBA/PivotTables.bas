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
Dim Cell As Range
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

For Each Cell In Range(ActiveCell.Address, ActiveCell.End(xlDown))

If Cell.IndentLevel > 0 Then

    Do While Cell.IndentLevel - counter > 0
        Cell.Insert Shift:=xlToRight
    counter = counter + 1
    Loop
    counter = 0
    
Else

Range(Cell.Offset(0, 1), Cell.Offset(0, 10)).ClearContents

End If

Next Cell

For Each Cell In Intersect(Columns(ColumnCount), ActiveSheet.UsedRange)
If IsEmpty(Cell.Offset(0, -1)) Then

FillCells = 1
Do Until FillCells = ColumnCount

Cells(Cell.Row, FillCells) = Cells(Cell.Row, FillCells).End(xlUp)
FillCells = FillCells + 1

Loop

Else

'Rows(Cell.Row).Delete

End If
Next Cell


End Sub

Sub test()
Dim pf As PivotField
Dim pt As PivotTable
Dim ColumnCount As Integer
Dim Cell As Range
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

For Each Cell In Range(Range("C3").Offset(1, ColumnCount), ActiveSheet.UsedRange.SpecialCells(xlCellTypeLastCell).Address)

'If the cell in column 1 on the corresponding row is a data point
If Cells(Cell.Row, 1).IndentLevel = ColumnCount - 1 Then

For IndentCount = 1 To ColumnCount

Set CellValue = Cells(Cell.Row, 1)
Do Until CellValue.IndentLevel = IndentCount - 1
Set CellValue = CellValue.Offset(-1, 0)
Loop
Cell.Offset(0, -ColumnCount + IndentCount - 1) = CellValue

Next IndentCount

Cell = Cell.Offset(0, -ColumnCount - 1)

Else

End If

Next Cell

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

