Attribute VB_Name = "generateManagerReport"
Option Explicit

Sub generateManagerReport()
Dim managerName As sliceritem
Dim managerNames() As String
Dim managersInReport As Integer

ReDim Preserve managerNames(0)
        
For Each managerName In ActiveWorkbook.SlicerCaches("Slicer_Manager_Name").SlicerItems

    If Not managerName.Selected Then
    
        If managerNames(0) <> "" Then ReDim Preserve managerNames(UBound(managerNames) + 1)
        
            managerNames(UBound(managerNames)) = managerName.Name
          
        Else
        
            managersInReport = managersInReport + 1
          
    End If
    
    If managersInReport > 10 Then
    
        MsgBox "A maximum of 10 managers can be selected for this report." & vbNewLine & _
        "Please choose 10 or fewer managers in the slicer above the " & Chr(34) & _
        "Generate Manager Specific Report" & Chr(34) & " button and run again."
        
        Exit Sub
    
    End If

Next
    
    'Filter the data set based on the managers listed
    Sheets("Data").ListObjects("Data").Range.AutoFilter Field:=12, Criteria1:= _
    managerNames, Operator:=xlFilterValues
    
    'TODO: Actually generate the specific report for the manager
        
        'Delete data that's not for the business unit
        Sheets("Data").Range("Data").Select
        Selection.EntireRow.Delete
        Sheets("Data").Range("A1").Select
        Selection.AutoFilter
        
    Application.ScreenUpdating = True
        'Refresh the pivot tables and charts based on the new data
        Sheets(1).Select
        ActiveSheet.Rows(1).Delete
        ActiveSheet.Rows(1).Delete
        ActiveWorkbook.RefreshAll
        
    Application.ScreenUpdating = False
        
            'Save the workbook with a business unit specific name
            ActiveWorkbook.SaveAs Filename:=parsedDashboardPath, FileFormat:=xlOpenXMLWorkbookMacroEnabled, CreateBackup:=False
            
    
    'Unfilter the data set on the Data worksheet.
    'Sheets("Data").ListObjects("Data").Range.AutoFilter
    

End Sub
