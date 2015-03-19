Attribute VB_Name = "Presentation"
Option Explicit

Sub CreateLegend() 'control As IRibbonControl
Dim NewSquare As Integer
Dim w As Integer
Dim myColor As String
Dim LegendCount As Integer

If ActiveSheet.ChartObjects.Count < 1 Then
LegendCount = InputBox("How many legend keys would you like?")
Else
LegendCount = ActiveSheet.ChartObjects(1).Chart.SeriesCollection.Count
End If

w = (ActiveSheet.UsedRange.SpecialCells(xlCellTypeLastCell).Row * 15) - (LegendCount * 15) - 10
myColor = 4

For NewSquare = 1 To LegendCount
w = w + 15
    ActiveSheet.Shapes.AddShape(msoShapeRectangle, Columns(1).Width - 10, w, 5, 5).Select
    Selection.ShapeRange.Line.Visible = msoFalse
    myColor = myColor + 1
    Selection.ShapeRange.Fill.ForeColor.ObjectThemeColor = myColor
    SendKeys "{ESC}"
    
Next NewSquare

'    Range("A1").Select
'    ActiveSheet.Shapes.SelectAll
'    Selection.Delete
        
End Sub
