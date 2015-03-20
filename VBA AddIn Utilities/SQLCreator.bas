Attribute VB_Name = "SQLCreator"
Option Explicit

Sub CreateSQLTable(control As IRibbonControl)
Dim Header As Range
Dim TableSQL As String
Dim TableName As String
Dim DataType As String
Dim clipboard As MSForms.DataObject
Set clipboard = New MSForms.DataObject

If UCase(Left(ActiveSheet.Name, 5)) <> "SHEET" And UCase(Left(ActiveSheet.Name, 5)) <> "NEW_U" Then
    TableName = Left(Replace(ActiveSheet.Name, " ", "_"), 30)
Else
TableName = Left(Replace(InputBox("What the name of your table?"), " ", "_"), 30)
ActiveSheet.Name = TableName
End If

TableSQL = "CREATE TABLE " & TableName & " ("

For Each Header In Range("A1:" & Range("A1").End(xlToRight).Address)

    If IsDate(Range(Header.Offset(1, 0).Address)) Then
        DataType = " Date"
    ElseIf WorksheetFunction.IsNumber(Header.Offset(1, 0)) Then
        DataType = " Number (10)"
    Else
        DataType = " VarChar(50)"
    End If

If Header = "Date" Then Header = "DUR"

'If UCase(Right(Header, 2)) = "ID" And Header.Column = 1 Then
'TableSQL = TableSQL & vbNewLine & "         " & Replace(Header, " ", "_") & DataType & " PRIMARY KEY" & ","
'Else
TableSQL = TableSQL & vbNewLine & "         " & Replace(Header, " ", "_") & DataType & ","
'End If
Next

TableSQL = Left(TableSQL, Len(TableSQL) - 1) & ");"

clipboard.SetText TableSQL
clipboard.PutInClipboard

Call UsageLog("Instert SQL Generator")

End Sub

Sub InsertSQL(control As IRibbonControl) '

Dim WorkRange As Range
Dim InsertSQL As String
Dim TableName As String
Dim Cell As Range
Dim Header As Range
Dim Values As String
Dim Continue As String
Dim Fields As String
Dim IncludeFieldNames As String
Dim clipboard As MSForms.DataObject
Set clipboard = New MSForms.DataObject

If IsEmpty(Range("A3")) Then Exit Sub

'Check to ensure each cell is filled in the current region.
If WorksheetFunction.CountA(Range("A1").CurrentRegion) <> Range("A1").CurrentRegion.Count Then
Continue = MsgBox("There are missing values in this list." & vbNewLine & vbNewLine & _
        "Do you want to continue?", vbYesNo)
    If Continue = vbNo Then Exit Sub
End If

IncludeFieldNames = MsgBox("Include field names?", vbYesNo)

If UCase(Left(ActiveSheet.Name, 5)) <> "SHEET" And UCase(Left(ActiveSheet.Name, 5)) <> "NEW_U" Then
    TableName = Replace(ActiveSheet.Name, " ", "_")
Else
TableName = Replace(InputBox("What table is this data being input into?"), " ", "_")
ActiveSheet.Name = TableName
End If

For Each Cell In Range("A1:" & Range("A1").End(xlToRight).Address)
Fields = Fields & Replace(Cell, " ", "_") & ", "
Next Cell
Fields = Left(Fields, Len(Fields) - 2)

For Each Cell In Range("A2:A" & Range("A2").End(xlDown).Row)

For Each Header In Range("A1:" & Range("A1").End(xlToRight).Address)
If Not IsEmpty(Cell.Offset(0, Header.Column - 1)) Then
    If IsDate(Range(Cell.Offset(0, Header.Column - 1).Address)) Then
        Values = Values & "TO_DATE('" & Format(Cell.Offset(0, Header.Column - 1), "yyyy-mm-dd") & "','YYYY-MM-DD'), "
    ElseIf WorksheetFunction.IsNumber(Cell.Offset(0, Header.Column - 1)) Then
        Values = Values & Cell.Offset(0, Header.Column - 1) & ", "
    Else
        Values = Values & "'" & Replace(Cell.Offset(0, Header.Column - 1), "'", "''") & "', "
    End If
Else
    Values = Values & "NULL" & ", "

End If

Next Header

Values = Left(Values, Len(Values) - 2)

If IncludeFieldNames = vbYes Then
InsertSQL = InsertSQL & vbNewLine & "INSERT INTO " & TableName & " (" & Fields & ") VALUES (" & Values & ");"
Else
InsertSQL = InsertSQL & vbNewLine & "INSERT INTO " & TableName & " VALUES (" & Values & ");"

End If

Values = ""

Next Cell

'InsertSQL = InsertSQL & vbNewLine & vbNewLine & "SELECT * FROM " & TableName & ";"

clipboard.SetText InsertSQL
clipboard.PutInClipboard

Call UsageLog("Instert SQL Generator")

End Sub

Sub CreateDecoder(control As IRibbonControl)
'SQL

Dim WorkRange As Range
Dim Cell As Range
Dim Expression As String
Dim FieldName As String
Dim IncludeCode As String
Dim clipboard As MSForms.DataObject
Set clipboard = New MSForms.DataObject

If Selection.Cells.Count > 200 Then Exit Sub
If Selection.Columns.Count <> 2 Then Exit Sub
If Selection.Areas.Count <> 1 Then Exit Sub

IncludeCode = MsgBox("Do you want to include the field value in the decoder?", vbYesNo)

FieldName = InputBox("Please enter the field name to be deconded.", "Field Name")
    
    Expression = "CASE " & FieldName & vbNewLine
    
For Each Cell In Selection

Cell = Replace(Cell, "'", "")
Cell = Replace(Cell, "--", "-")

If Left(Cell.Address, 2) = Left(Selection.Columns(1).Address, 2) Then

    If WorksheetFunction.IsText(Cell) Then
        Expression = Expression & "  WHEN '" & Cell & "' "
        Else
        Expression = Expression & "  WHEN " & Cell & " "
    End If
    
    Else

    If WorksheetFunction.IsText(Cell) Then
    
            If IncludeCode = vbYes Then
                Expression = Expression & " THEN '" & Cell.Offset(0, -1) & " - " & Cell & "'" & vbNewLine
            Else
                Expression = Expression & " THEN '" & Cell & "'" & vbNewLine
            End If
    
        Else
    
            If IncludeCode = vbYes Then
                Expression = Expression & " THEN '" & Cell.Offset(0, -1) & " - " & Cell & "" & vbNewLine
            Else
                Expression = Expression & " THEN '" & Cell & "'" & vbNewLine
            End If
        
    End If
    
End If

Next Cell

Expression = Expression & "  ELSE " & FieldName & vbNewLine & "END"

clipboard.SetText Expression
clipboard.PutInClipboard

Call UsageLog("SQL Decoder")

End Sub

Sub ToDate(control As IRibbonControl)
Dim FieldName As String
Dim clipboard As MSForms.DataObject
Set clipboard = New MSForms.DataObject



If IsEmpty(ActiveCell) Then

clipboard.GetFromClipboard
FieldName = "TO_DATE(" & clipboard.GetText(1) & ",'YYYY-MM-DD')"

Else

If IsDate(ActiveCell) Then
    If ActiveCell <> Int(ActiveCell) Then
        FieldName = "TO_DATE('" & ActiveCell & "','MM-DD-YYYY HH:MI:SS AM')"
    Else
        FieldName = "TO_DATE('" & ActiveCell & "','MM-DD-YYYY')"
    End If
Else
    MsgBox "Please select a cell with a date. Or select a blank cell to use the content on your clipboard for the expression."
    Exit Sub
End If
End If

If InStr(FieldName, "(,") <> 0 Then Exit Sub

clipboard.SetText FieldName
clipboard.PutInClipboard

Debug.Print FieldName

End Sub

Sub CreateSelect(control As IRibbonControl)
Dim Header As Range
Dim SelectSQL As String
Dim TableName As String
Dim clipboard As MSForms.DataObject
Set clipboard = New MSForms.DataObject

If UCase(Left(ActiveSheet.Name, 5)) <> "SHEET" And UCase(Left(ActiveSheet.Name, 5)) <> "NEW_U" Then
    TableName = Left(Replace(ActiveSheet.Name, " ", "_"), 30)
Else
TableName = Left(Replace(InputBox("What is the name of your table?"), " ", "_"), 30)
ActiveSheet.Name = TableName
End If

For Each Header In Range("A1:" & Range("A1").End(xlToRight).Address)

SelectSQL = SelectSQL & Replace(Header, " ", "_") & ", "

Next Header

SelectSQL = "SELECT " & Left(SelectSQL, Len(SelectSQL) - 2) & " FROM " & TableName & ";"

clipboard.SetText SelectSQL
clipboard.PutInClipboard

Call UsageLog("Select SQL Generator")

End Sub
