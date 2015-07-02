Attribute VB_Name = "FormatReports"
Option Explicit
'These macros format reports exported from PS

Sub GoldsGymReport(control As IRibbonControl)
Dim lastRow As Integer, cell As Range, TextRows As String
Dim Continue As String, PayDay As Date
Dim i As Integer

If InStr(1, Range("B2"), "PAY001") = 0 Then
    Continue = MsgBox("The active sheet does not appear to be the " & _
        "Gold's Gym Report." & vbNewLine & vbNewLine & "Do you want to continue?", vbYesNo + vbDefaultButton2 + vbQuestion, "Continue?")
    If Continue = vbNo Then GoTo EndEarly
End If

Application.ScreenUpdating = False

PayDay = Range("B4") + 7

Set cell = Range("A1")

Do Until InStr(1, cell, "Employee ID") <> 0
    Set cell = cell.Offset(1, 0)
Loop

Rows("1:" & cell.Offset(-1, 0).Row).Delete

lastRow = ActiveSheet.UsedRange.SpecialCells(xlCellTypeLastCell).Row

For i = lastRow To 2 Step -1
    If Range("A" & i).Value = "" Then
        Range(i & ":" & i).Delete
    End If
Next i

Columns("C").Delete
Columns("D:O").Delete

Rows("1:2").Insert

Columns("A").Select

    With Range("A1").Font
        .Name = "Calibri"
        .Size = 16
        .Bold = True
    End With
    With Range("A1:D1").Borders(xlEdgeBottom)
        .Weight = xlMedium
    End With
    
    With Range("A3:D3").Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0.599993896298105
        .PatternTintAndShade = 0
    End With
    With Range("A3:D3").Borders(xlEdgeBottom)
        .Weight = xlMedium
    End With
    
Range("A1") = "LDS Church Payroll - Vasa Fitness Gym Deductions for Payroll Date " & PayDay
Range("B3") = "Employee Name"
Range("C3") = "Deduction Amount"
Range("D3") = "Comments"

lastRow = Range("A3").End(xlDown).Row

For Each cell In Range("D4:D" & lastRow)
    If cell.Offset(0, -1) = 0 Then cell = "No payroll for this pay period"
Next

   Columns("A:A").ColumnWidth = 12
    Columns("C:C").NumberFormat = "$#,##0.00_);[Red]($#,##0.00)"
    Columns("C:C").ColumnWidth = 18
    Columns("D:D").ColumnWidth = 40
    Rows("3:3").EntireRow.AutoFit

    
ActiveSheet.Name = Replace(PayDay, "/", "-")
ActiveWorkbook.SaveAs fileName:=ActiveWorkbook.Path & "\Gold's Gym Report " & Replace(PayDay, "/", "-"), FileFormat:=51

MsgBox "The Gold's Gym report for paydate " & PayDay & " has been completed." & vbNewLine & vbNewLine & _
        "Please make sure to update the notes in column C before sending to ..."

EndEarly:

Range("A1").Select

Application.ScreenUpdating = True

End Sub

Sub FormatPSALP(control As IRibbonControl)
Dim lastRow As Integer
Dim LastCell As String
Dim cell As Range
Dim SendEmpIDs As Range
Dim Continue As String

If InStr(Range("B2"), "CPAY002I") = 0 Then
    Continue = MsgBox("The active sheet does not appear to be the " & _
        "ALP report." & vbNewLine & vbNewLine & "Do you want to continue?", vbYesNo + vbDefaultButton2 + vbQuestion, "Continue?")
    If Continue = vbNo Then Exit Sub
End If

ActiveSheet.Name = "ALP PP" & RecentPP()

Application.ScreenUpdating = False

lastRow = ActiveSheet.UsedRange.SpecialCells(xlCellTypeLastCell).Row

    Columns("A:A").AutoFilter
    ActiveSheet.Range("$A:$A").AutoFilter Field:=1, Criteria1:="=Name", _
        Operator:=xlAnd
    Range("A6:A" & lastRow).SpecialCells(xlCellTypeVisible).EntireRow.Delete
    Columns("A:A").AutoFilter
    
    Columns("C:C").SpecialCells(xlCellTypeBlanks).EntireRow.Delete
    Rows("1:1").Delete Shift:=xlUp
    
LastCell = ActiveSheet.UsedRange.SpecialCells(xlCellTypeLastCell).Offset(0, -1).Address
    
    With Range("A3:" & LastCell)
        .SpecialCells(xlCellTypeBlanks).FormulaR1C1 = "=R[-1]C"
        .Value = .Value
    End With
    
lastRow = Range("A1").End(xlDown).Row
    
For Each cell In Range("F2:F" & lastRow)
    cell.Value = cell.Value & cell.Offset(0, 1)
Next cell

Range("G2:G" & lastRow).Delete Shift:=xlToLeft
        
For Each cell In Range("D2:D" & lastRow)
    If InStr(cell, "-") Then
        cell = -Left(cell, Len(cell) - 1)
    End If
Next cell

Range("C1") = "LMP Hours"
Range("D1") = "LML Hours"

Columns(12).Insert

Range("L1") = "LMP  Positive Input"
Range("M1") = "LML  Positive Input"

For Each cell In Range("M2:M" & lastRow)
    If InStr(cell, "-") Then
        cell = -Left(cell, Len(cell) - 1)
    End If
Next cell

For Each cell In Range("M2:M" & lastRow)
        
    If cell.Offset(0, -11) = cell.Offset(1, -11) Then
         
        cell.Offset(1, 1).Cut Destination:=cell.Offset(0, 2)
        cell.Offset(1, 0).Cut Destination:=cell.Offset(0, -1)
        cell.Offset(1, -9).Cut Destination:=cell.Offset(0, -10)
        cell.Offset(1, 0).EntireRow.Delete
            
    End If

Next cell

lastRow = Range("A1").End(xlDown).Row

'Range("M2:N" & LastRow).Replace What:="***", Replacement:=", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False

For Each cell In Range("N2:N" & lastRow)
    
    If Not IsEmpty(cell) And Not IsEmpty(cell.Offset(0, 1)) Then
        cell = Right(cell, Len(cell) - 4) & " - " & Right(cell.Offset(0, 1), Len(cell.Offset(0, 1)) - 4)
    End If
    
If IsEmpty(cell) And Not IsEmpty(cell.Offset(0, 1)) Then cell = Right(cell.Offset(0, 1), Len(cell.Offset(0, 1)) - 4)

If Not IsEmpty(cell) And IsEmpty(cell.Offset(0, 1)) Then cell = Right(cell, Len(cell) - 4)
    
Next cell

Range("O2:O" & lastRow).ClearContents

With Range("N2:N" & lastRow)
    .Replace What:="Bal", Replacement:="balance", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    .Replace What:="      ", Replacement:=" ", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    .Replace What:="     ", Replacement:=" ", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    .Replace What:="    ", Replacement:=" ", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    .Replace What:="PI", Replacement:="AM Positive Input", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
End With

'Generate description of Payroll Status
Range("I1") = "Payroll Status"

    Range("I:I").Replace What:="A", Replacement:="Active", LookAt:=xlWhole, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Range("I:I").Replace What:="D", Replacement:="Deceased", LookAt:=xlWhole, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Range("I:I").Replace What:="L", Replacement:="Leave of Absence", LookAt:=xlWhole, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Range("I:I").Replace What:="P", Replacement:="Leave With Pay", LookAt:=xlWhole, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Range("I:I").Replace What:="Q", Replacement:="Retired With Pay", LookAt:=xlWhole, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Range("I:I").Replace What:="R", Replacement:="Retired", LookAt:=xlWhole, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Range("I:I").Replace What:="S", Replacement:="Suspended", LookAt:=xlWhole, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Range("I:I").Replace What:="T", Replacement:="Terminated", LookAt:=xlWhole, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Range("I:I").Replace What:="U", Replacement:="Terminated With Pay", LookAt:=xlWhole, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Range("I:I").Replace What:="V", Replacement:="Terminated Pension Pay Out", LookAt:=xlWhole, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Range("I:I").Replace What:="W", Replacement:="Short Work Break", LookAt:=xlWhole, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Range("I:I").Replace What:="X", Replacement:="Retired-Pension Administration", LookAt:=xlWhole, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False


Range("O1") = "Action Needed"

'Provide some detail to the messages in the report.
For Each cell In Range("O2:O" & lastRow)

    If InStr(cell.Offset(0, -1), "Paysheet differs from AM Positive Input.") Then
        cell = "Check employee job data status and ensure balance adjustment has been made."
    ElseIf InStr(cell.Offset(0, -1), "Off Cycle Payout indicated.") Then
        cell = "Off-Cycle check already processed for ALP. If an on-cycle payment for this employee is also listed on this report, you may need to uncheck to OK to Pay."
    ElseIf InStr(cell.Offset(0, -1), "Previous Payout") Then
        cell = "Ensure previous hours aren't paid out a second time."
    ElseIf InStr(cell.Offset(0, -1), "Multiple Vacation Payouts.") Then
        cell = "Ensure multiple payouts are not duplicate payments."
    End If
    
Next cell

    Rows("1:1").Font.Bold = True
    Range("A1").Select
    
    Columns("C:E").Style = "Comma"
    Columns("G:H").NumberFormat = "m/d/yyyy"
    Columns("L:M").Style = "Comma"
    
Columns.AutoFit

Columns("N:N").AutoFilter

Application.ScreenUpdating = True

If MsgBox("Save to the shared drive?", vbYesNo) = vbYes Then
    Application.ActiveWorkbook.SaveAs fileName:="\\CHQPVUN0066\FINUSR\SHARED\FIN_PYRL\2_Payroll Time & Labor Absence Management\Processed (Historic)\ALP Comparisons\" & ActiveSheet.Name & ".xlsx", FileFormat:=51
End If

End Sub

Sub OnlineCheckReportwithTRC(control As IRibbonControl)
Dim List As Range, CompareRow As Integer, lastRow As Integer, ActiveRow As Range, cell As Range
Dim EmployeeCount As Integer, Continue As String

If InStr(1, Range("B2"), "CPAY544A") = 0 Then
    Continue = MsgBox("The active sheet does not appear to be the " & _
        "Online Check report." & vbNewLine & vbNewLine & "Do you want to continue?", vbYesNo + vbDefaultButton2 + vbQuestion, "Continue?")
    If Continue = vbNo Then Exit Sub
End If

ActiveSheet.Name = "PP" & InputBox("Please enter the two digit pay period.", "Pay Period")

Application.ScreenUpdating = False

    lastRow = Range("A5").End(xlDown).Row
    Rows(lastRow).Delete
    Rows(lastRow - 1).Delete
    
    Columns("D:F").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("D5") = "Earns Code"
    Range("E5") = "Hours"
    Range("f5") = "Amounts"
    CompareRow = 6
    Rows(5).Font.Bold = True

Application.DisplayAlerts = False
    ActiveSheet.Range("$A$1:$A$" & lastRow).AutoFilter Field:=1, Criteria1:="------"
    Range("A6:A" & lastRow).SpecialCells(xlCellTypeVisible).Rows.Delete
Application.DisplayAlerts = True
    
    Selection.AutoFilter
    
Do While Not IsEmpty(Range("A" & CompareRow))

    Do Until IsEmpty(Range("G" & CompareRow))
        CompareRow = CompareRow + 1
    Loop
    
    If IsEmpty(Range("D" & CompareRow - 1)) Then
        Range("A" & CompareRow & ":C" & CompareRow).Cut
        Range("D" & CompareRow - 1 & ":F" & CompareRow - 1).Select
        ActiveSheet.Paste
        Rows(CompareRow).Delete
        
        Do Until Not IsEmpty(Range("G" & CompareRow))
        If IsEmpty(Range("A" & CompareRow)) Then
        Exit Do
        End If

            Range("A" & CompareRow & ":C" & CompareRow).Cut
            Range("D" & CompareRow & ":F" & CompareRow).Select
            ActiveSheet.Paste
            
            Range("A" & CompareRow - 1 & ":C" & CompareRow - 1).Copy
            Range("A" & CompareRow & ":C" & CompareRow).PasteSpecial
            ActiveSheet.Paste
            
            Range("G" & CompareRow - 1 & ":R" & CompareRow - 1).Copy
            Range("G" & CompareRow & ":R" & CompareRow).PasteSpecial
            ActiveSheet.Paste
        Loop
        
        Else
        
        If Range("A" & CompareRow) <> "EMPLOYEE COUNT" Then
        
            Do Until Not IsEmpty(Range("G" & CompareRow))
                Range("A" & CompareRow & ":C" & CompareRow).Cut
                Range("D" & CompareRow & ":F" & CompareRow).Select
                ActiveSheet.Paste
                
                Range("A" & CompareRow - 1 & ":C" & CompareRow - 1).Copy
                Range("A" & CompareRow & ":C" & CompareRow).PasteSpecial
                ActiveSheet.Paste
                
                Range("G" & CompareRow - 1 & ":R" & CompareRow - 1).Copy
                Range("G" & CompareRow & ":R" & CompareRow).PasteSpecial
                ActiveSheet.Paste
            Loop
        
        End If
    
    End If

Loop

Rows("1:4").Delete
Range("R1").ClearContents

Range("A1").Select

Application.ScreenUpdating = True

End Sub

Sub SummarizeTRCReport(control As IRibbonControl)
Dim Cancel As String
Dim TRClastRow As String
Dim A7Name As String
Dim Continue As String

Application.ScreenUpdating = False

'Check for the Report Name in cell B2
If InStr(1, Range("B2"), "CTL908") = 0 Then
    Continue = MsgBox("The active sheet does not appear to be a " & _
        "TRC report." & vbNewLine & vbNewLine & "Do you want to continue?", vbYesNo + vbDefaultButton2 + vbQuestion, "Continue?")
    If Continue = vbNo Then Exit Sub
End If

    If Range("A7") = "NOTHING TO REPORT" Then
        
        MsgBox ("It appears the TRC report for this store is blank." & vbNewLine & vbNewLine & _
        "You may have forgotten to specify a DeptID when you ran the report or the time " & _
        "may not be loaded into PeopleSoft yet." & vbNewLine & vbNewLine & _
        "Please correct the report on sheet 2 and rerun the comparison.")
        Cancel = "Yes"
        Application.ScreenUpdating = True
        Exit Sub
    End If

    TRClastRow = Mid(ActiveSheet.UsedRange.Address, 9)

'Put the Name on the Same line as the EmpID and total hours.
    A7Name = Range("A7")
    
    Range("A6").Delete Shift:=xlUp
    Range("A7").Delete Shift:=xlUp
    
    ActiveSheet.Range("$C:$C").AutoFilter Field:=1, Criteria1:=">=0"

    Range("A1:A" & TRClastRow).Copy
    Range("K1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("C1:C" & TRClastRow).Copy
    Range("L1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("E1:E" & TRClastRow).Copy
    Range("M1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Selection.AutoFilter
    
    'Copy Header Values
        Range("D4").Copy Range("K1")
        Range("D5").Copy Range("L1")
        Range("B4").Copy Range("M1")
    Columns("A:J").Delete Shift:=xlToLeft
    
    'Add first employee name
    Range("A2").Insert (xlDown)
    Range("A2") = A7Name

    Range("A1:A" & TRClastRow).SpecialCells(xlCellTypeVisible).Select
    
    Rows("1:3").Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A4:C4").Cut Range("A3:C3")
    Range("A1:C3").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
    End With
    Range("A1:C1").Merge
    ActiveCell.FormulaR1C1 = "Employee Time by TRC"
    Range("A2:C2").Merge
    Range("A2:C2").FormulaR1C1 = "PeopleSoft"
    Range("A4").FormulaR1C1 = "Name"
    Range("B4").FormulaR1C1 = "EmpID"
    Range("C4").FormulaR1C1 = "Total Hours"
    
With Range("A4:C4")
    .Font.Bold = True
End With

Application.DisplayAlerts = False
    Rows(1).Delete
    Rows(1).Delete
Application.DisplayAlerts = True

    TRClastRow = Range("A2").End(xlDown).Row

Range("C3:C" & TRClastRow).NumberFormat = "0.00"

Range("A1").Select

Application.ScreenUpdating = True

End Sub

Sub CTL916(control As IRibbonControl)

    Dim FileDate As String

    Selection.AutoFilter
    ActiveSheet.Range(ActiveCell.CurrentRegion.Address).AutoFilter Field:=1, Criteria1:="------"
    ActiveCell.CurrentRegion.Select
    Selection.SpecialCells(xlCellTypeVisible).Select
    Selection.EntireRow.Delete
    Range("D5").FormulaR1C1 = "TRC Category"
    Range("E5").FormulaR1C1 = "Reported Quanity"
    Range("F5").FormulaR1C1 = "Payable Quantity"
    Range("G5").FormulaR1C1 = "TaskGroup"
    Range("H5").FormulaR1C1 = "WorkGroup"
    Range("I5").FormulaR1C1 = "Status"
    Range("J5").FormulaR1C1 = "ECD"
    Range("K5").FormulaR1C1 = "Action Date"
    Range("L5").FormulaR1C1 = "Message Data"
    Range("A5").FormulaR1C1 = "EmpID"
    Range("B5").FormulaR1C1 = "EmplRcd"
    Range("C5").FormulaR1C1 = "Date Reported"
    Rows("1:4").Delete Shift:=xlUp
    
    Cells.Select
    Cells.EntireColumn.AutoFit
    Columns("A:A").ColumnWidth = 8.57
    
    Range("A1").Select
    
    If (Len(Month(Date)) = 1) Then
        FileDate = "0" & Month(Date)
    Else
        FileDate = Month(Date)
    End If
    
    If (Len(Day(Date)) = 1) Then
        FileDate = FileDate & "0" & Day(Date)
    Else
        FileDate = FileDate & Day(Date)
    End If
    
    FileDate = FileDate & Year(Date)
    
    Application.ActiveWorkbook.SaveAs fileName:="\\CHQPVUN0066\FINUSR\SHARED\FIN_PYRL\2_Payroll Time " & _
        "& Labor Absence Management\Processed (Historic)\CTL916\CTL916 " & FileDate & ".xlsx", FileFormat:=51
    
End Sub
