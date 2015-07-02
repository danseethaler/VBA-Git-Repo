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

Dim parameter As IRibbonControl
Call ColumnsAutofit(parameter)
    
Range("A1").Select

Application.ScreenUpdating = True

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

End Sub

Sub ListCriteria(control As IRibbonControl)

Dim convertEmpIDs As String
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
        
        'Determine if the cell is not in a PS EmpID format.
        'If not, ask the user if the cell should be converted to an EmpID format
        If Len(cell) < 6 Or cell.NumberFormat <> "@" Then
            
            If convertEmpIDs = "" Then
                convertEmpIDs = MsgBox("Convert EmpIDs to text?", vbYesNo)
            End If
            
            If convertEmpIDs = vbYes Then cell = formatEmpID(cell)
        
        End If
    
        If Len(List & cell.Value) <= 255 Then
        
            List = List & cell.Value & "','"
        
        Else
        
            Exit For
            
        End If
    
    End If
    
    cell.Offset(1, 0).Select

Next cell

'Truncate the ',' from the end of the list
List = Left(List, Len(List) - 3)

'If we've reached the end of the used range, send the active cell to the top of the worksheet
If Intersect(ActiveCell, ActiveSheet.UsedRange) Is Nothing Then
    ActiveCell.Offset(-1, 0).End(xlUp).Select
End If

clipboard.SetText List
clipboard.PutInClipboard

Application.ScreenUpdating = True

MsgBox ("This list has been copied to your clipboard." & vbNewLine & vbNewLine & _
        "Paste this list in the 'List Members' of a PeopleSoft Query criteria for bulk processing.")

End Sub

Private Function formatEmpID(cell As Range) As Range

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

Set formatEmpID = cell

End Function

Sub ReopenUnsavedWorkbook(control As IRibbonControl)
    Dim ConfirmSub
    Dim openWorkbook As String

    openWorkbook = ActiveWorkbook.FullName

    Application.DisplayAlerts = False
        ActiveWorkbook.Close (False)
    Application.DisplayAlerts = True
    
    Application.Workbooks.Open (openWorkbook)

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

'MsgBox NameCounter & " name(s) were reformatted to 'First Last' format."

End Sub

Sub FormulastoValues(control As IRibbonControl)
Dim workRange As Range
Dim intArea As Integer
Dim rngCell As Range
    
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

'MsgBox (WorkRange.Cells.Count & " formula(s) converted to values.")

End Sub

Sub CreateDecoder(control As IRibbonControl)
'SQL - Create a decoder based on the

Dim workRange As Range
Dim cell As Range
Dim Expression As String
Dim FieldName As String
Dim IncludeCode As String
Dim clipboard As MSForms.DataObject
Set clipboard = New MSForms.DataObject

If Selection.Cells.Count > 200 Then Msgobox "Too many cells selected.": Exit Sub
If Selection.Columns.Count <> 2 Then Msgobox "Only two columns can be selected.":  Exit Sub
If Selection.Areas.Count <> 1 Then Msgobox "Only one area can be selected.":  Exit Sub

IncludeCode = MsgBox("Do you want to include the field value in the decoder?", vbYesNo)

FieldName = InputBox("Please enter the field name to be deconded.", "Field Name")
    
    Expression = "CASE " & FieldName & vbNewLine
    
For Each cell In Selection

cell = Replace(cell, "'", "")
cell = Replace(cell, "--", "-")

If Left(cell.Address, 2) = Left(Selection.Columns(1).Address, 2) Then

    If WorksheetFunction.IsText(cell) Then
        Expression = Expression & "  WHEN '" & cell & "' "
        Else
        Expression = Expression & "  WHEN " & cell & " "
    End If
    
    Else

    If WorksheetFunction.IsText(cell) Then
    
            If IncludeCode = vbYes Then
                Expression = Expression & " THEN '" & cell.Offset(0, -1) & " - " & cell & "'" & vbNewLine
            Else
                Expression = Expression & " THEN '" & cell & "'" & vbNewLine
            End If
    
        Else
    
            If IncludeCode = vbYes Then
                Expression = Expression & " THEN '" & cell.Offset(0, -1) & " - " & cell & "" & vbNewLine
            Else
                Expression = Expression & " THEN '" & cell & "'" & vbNewLine
            End If
        
    End If
    
End If

Next cell

Expression = Expression & "  ELSE " & FieldName & vbNewLine & "END"

clipboard.SetText Expression
clipboard.PutInClipboard

End Sub

Sub RemoveLeadingSpaces(control As IRibbonControl)
Dim cell As Range

For Each cell In Selection
        Do Until Left(cell, 1) <> " "
            cell = Right(cell, Len(cell) - 1)
        Loop
Next cell

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
