Attribute VB_Name = "CalledSubs"
Option Explicit

'This subroutine writes a line to the usage log with the user's
'username, subroutine performed, and when they used it.

Sub UsageLog(SubName As String)
Dim directoryPath As String
On Error Resume Next

If Environ("username") = "danielseethaler" Or Environ("username") = "danseethaler" Then Exit Sub

    directoryPath = "\\CHQPVUN0066\FINUSR\SHARED\FIN_PYRL\2_Payroll Time & Labor Absence Management\Desk Manual (Information)\"

    Open directoryPath & "Usage Log.txt" For Append As #1
        If Err <> 0 Then GoTo NextSub
    Print #1, SubName & "," & Environ("username") & "," & Now
    Close #1
    
NextSub:
    
    Err = 0
    
    directoryPath = "\\N0053_SERVER\GSCUSR\SHARED\HR Processing Services\Utilities\"

    Open directoryPath & "Usage Log.txt" For Append As #1
        If Err <> 0 Then Exit Sub
    Print #1, SubName & "," & Environ("username") & "," & Now
    Close #1

On Error GoTo 0

End Sub


Sub ReverseNameCall(FormatRange As Range)

Dim FirstName As String, LastName As String
Dim cell As Range
Dim workRange As Range

On Error Resume Next

    If FormatRange.Cells.Count = 1 Then
        If Not IsEmpty(ActiveCell) And Not ActiveCell.HasFormula Then
            
            Set workRange = FormatRange
        
        Else

            MsgBox "Please select a range with names in 'Last, First' format.", vbInformation
            Exit Sub
        
        End If
    
    ElseIf FormatRange.Cells.Count > 1 Then

        Set workRange = FormatRange.SpecialCells(xlCellTypeConstants)
        Set workRange = Intersect(workRange, workRange.Parent.UsedRange)
    
    End If
    
    If Err = 1004 Or Err = 94 Or Err = 91 Then
        MsgBox "There are no constants in your selection", vbInformation
        Exit Sub
    
    ElseIf Err = 6 Then
    
        Set workRange = FormatRange.SpecialCells(xlCellTypeConstants)
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
        
        End If
        
        
        If InStr(cell, "  ") Then
            
            LastName = Left(cell, InStr(cell, "  ") - 1)
            FirstName = Right(cell, Len(cell) - InStr(cell, "  "))
        
            If Left(FirstName, 1) = " " Then
                FirstName = Right(FirstName, Len(FirstName) - 1)
            End If
            
            cell = WorksheetFunction.Proper(FirstName & " " & LastName)
        
        End If
        
        Do Until Left(cell, 1) <> " "
            cell = Right(cell, Len(cell) - 1)
        Loop
    
    End If

Next cell

End Sub

Sub ConvertEmpIDToTextCall(EmpIDRange As Range)
Dim cell As Range
Dim workRange As Range

On Error Resume Next

    If EmpIDRange.Cells.Count = 1 Then
        If Not IsEmpty(ActiveCell) And Not ActiveCell.HasFormula Then
            
            Set workRange = EmpIDRange
        
        Else

            MsgBox "Please select a range with constants", vbInformation
            Exit Sub
        
        End If
    
    ElseIf EmpIDRange.Cells.Count > 1 Then

        Set workRange = EmpIDRange.SpecialCells(xlCellTypeConstants)
        Set workRange = Intersect(workRange, workRange.Parent.UsedRange)
    
    End If
    
    If Err = 1004 Or Err = 94 Or Err = 91 Then
        MsgBox "There are no constants in your EmpIDRange", vbInformation
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

For Each cell In workRange

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

End Sub

Sub ColumnsAutofitCall()
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


Sub VerifyReferenceLibraries()


    Dim theRef As Variant
    Dim i As Long
     
     'Set to continue in case of error
    On Error Resume Next
     
     'Clear any errors so that error trapping for GUID additions can be evaluated
    Err.Clear
     
     'Add the reference
        ThisWorkbook.VBProject.References.AddFromGuid GUID:="{000204EF-0000-0000-C000-000000000046}", Major:=1, Minor:=0
        ThisWorkbook.VBProject.References.AddFromGuid GUID:="{00020813-0000-0000-C000-000000000046}", Major:=1, Minor:=0
        ThisWorkbook.VBProject.References.AddFromGuid GUID:="{00020430-0000-0000-C000-000000000046}", Major:=1, Minor:=0
        ThisWorkbook.VBProject.References.AddFromGuid GUID:="{2DF8D04C-5BFA-101B-BDE5-00AA0044DE52}", Major:=1, Minor:=0
        ThisWorkbook.VBProject.References.AddFromGuid GUID:="{0D452EE1-E08F-101A-852E-02608C4D0BB4}", Major:=1, Minor:=0
        ThisWorkbook.VBProject.References.AddFromGuid GUID:="{00062FFF-0000-0000-C000-000000000046}", Major:=1, Minor:=0
        ThisWorkbook.VBProject.References.AddFromGuid GUID:="{0002E157-0000-0000-C000-000000000046}", Major:=1, Minor:=0
        ThisWorkbook.VBProject.References.AddFromGuid GUID:="{420B2830-E718-11CF-893D-00A0C9054228}", Major:=1, Minor:=0
     
     Call UsageLog("Reference Libraries Verified")
     
     'If an error was encountered, inform the user
    Select Case Err.Number
    Case Is = 32813
         'Reference already in use.  No action necessary
    Case Is = vbNullString
         'Reference added without issue
    Case Else
         'An unknown error was encountered, so alert the user
        MsgBox "A problem was encountered trying to" & vbNewLine _
        & "add or remove a reference in this file" & vbNewLine & "Please check the " _
        & "references in your VBA project!", vbCritical + vbOKOnly, "Error!"
    End Select
    On Error GoTo 0

End Sub
