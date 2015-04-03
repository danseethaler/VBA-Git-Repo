Attribute VB_Name = "DivideAndEmail"
Option Explicit

    Dim payPeriod As Integer
    Dim businessUnit As Range
    Dim parsedDashboardPath As String
    Dim masterDashboardPath As String
    Dim Outlook As Outlook.Application
    Dim MyItem As Outlook.MailItem

Sub parseDashboards()

Dim openWorkbook As Workbook
Dim emailControllers As String

'payPeriod = RecentPP()
payPeriod = 6

'Set the directory path for the master dashboard
masterDashboardPath = "\\CHQPVUN0066\FINUSR\SHARED\FIN_PYRL\2_Payroll Time & Labor Absence Management\Processed (Historic)\" & _
    "Reported Time Dashboard\Time Entry Dashboard - PP" & payPeriod & " " & Year(Date) & " Master.xlsm"

'Ensure the master file for the pay period we're processing exists.
If Dir(masterDashboardPath) = "" Then
    MsgBox "Please create the file " & masterDashboardPath & " before you proceed."
    Exit Sub
End If

'Make sure the master workbook isn't currently open
For Each openWorkbook In Application.Workbooks

    If openWorkbook.Name = "Time Entry Dashboard - PP" & payPeriod & " " & Year(Date) & " Master.xlsm" And openWorkbook.Saved = False Then
    
            MsgBox "Please close the master workbook before running this macro."
            Exit Sub
            
        ElseIf openWorkbook.Name = "Time Entry Dashboard - PP" & payPeriod & " " & Year(Date) & " Master.xlsm" Then
        
            openWorkbook.Close
            
    End If
    
Next openWorkbook

emailControllers = MsgBox("Would you like to email the files?", vbYesNoCancel)

If emailControllers = vbCancel Then Exit Sub

With Application
    .ScreenUpdating = False
    .EnableEvents = False
End With

        'Initialize Outlook
        If emailControllers = vbYes Then
            
            'Check to see if Outlook is open.
            On Error Resume Next
                Set Outlook = GetObject(, "Outlook.Application")
            On Error GoTo 0
            
            'If Outlook is not open then open it.
            If Outlook Is Nothing Then
                Set Outlook = CreateObject("Outlook.Application")
                Outlook.GetNamespace("MAPI").GetDefaultFolder(olFolderInbox).Display
            End If
            
        End If

'Iterate through each visible cell in the Business Unit column
For Each businessUnit In Range("DL[Business Unit]").SpecialCells(xlCellTypeVisible)

parsedDashboardPath = "\\CHQPVUN0066\FINUSR\SHARED\FIN_PYRL\2_Payroll Time & Labor Absence Management\" & _
    "Processed (Historic)\Reported Time Dashboard\PP" & payPeriod & "\Time Entry Dashboard - PP" & payPeriod & " " & businessUnit & ".xlsm"

    'If the process column is set to "Yes" and the file does not already exist for the pay period
    If Dir(parsedDashboardPath) = "" And Not IsEmpty(Intersect(Rows(businessUnit.Row), Range("DL[Controller Email]"))) Then
    
        'Open the master workbook
        Workbooks.Open Filename:=masterDashboardPath
        
            'Make sure the master workbook is active and select the data tab
            Windows("Time Entry Dashboard - PP" & payPeriod & " " & Year(Date) & " Master.xlsm").Activate
            Sheets("Data").Select
        
            'Filter the data based on the business unit we're processing
            Range("Data[#All]").AutoFilter Field:=3, Criteria1:= _
                "<>" & businessUnit, Operator:=xlAnd
            
            'Delete data that's not for the business unit
            Range("Data").Select
            Selection.EntireRow.Delete
            Range("A1").Select
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
            
            'Close the workbook
            Windows("Time Entry Dashboard - PP" & payPeriod & " " & businessUnit & ".xlsm").Close
        
        'If the user selected to email the controllers then call the emailDashboards macro
        If emailControllers = vbYes Then
        
            Call emailDashboards
        
        End If
        
        Windows("Time Entry Dashboard - VBA to Parse and Email.xlsm").Activate
    
    End If

Next

With Application
    .ScreenUpdating = True
    .EnableEvents = True
End With

End Sub

Sub emailDashboards()
'This macro sends the parsed file to the emails listed on the associated DL

    'This line creates the Outlook mail object and assigns it to the designated template.
    Set MyItem = Outlook.CreateItemFromTemplate("\\CHQPVUN0066\FINUSR\SHARED\FIN_PYRL\2_Payroll Time & Labor Absence Management\" & _
        "Processed (Historic)\Reported Time Dashboard\Time Approval Dashboard.oft")
    
    'This section manipulates several of the properties of the template to insert
    'the information on the row the program is processing.
    With MyItem
    
        .Attachments.Add parsedDashboardPath
        .Subject = "Reported Time Dashboard for PP" & payPeriod
        .To = Intersect(Rows(businessUnit.Row), Range("DL[Controller Email]")) & _
            ";" & Intersect(Rows(businessUnit.Row), Range("DL[HR Email]"))
        .CC = "danseethaler@ldschurch.org"
        .Subject = "Reported Time Dashboard for PP" & payPeriod & " - " & businessUnit
        .HTMLBody = Replace(MyItem.HTMLBody, "businessUnit", businessUnit)
        .HTMLBody = Replace(MyItem.HTMLBody, "payPeriod", payPeriod)
        .Display
        
    End With
    
End Sub

Public Function RecentPP() As Integer
'This function returns the most recently completed pay period number as an integer.

Dim firstPayPeriod As Date
Dim calculatedPayPeriod As Date
Dim lastPPEndDate As Date

'Determine the most recent PP end date based on a static date in the past
If (Date - CDate("8/15/2014")) Mod 14 = 0 Then

        lastPPEndDate = Date - 14
    
    Else
    
        lastPPEndDate = Date - (Date - CDate("8/15/2014")) Mod 14

End If

'Determine the PPEnd date for firstPayPeriod of the current year
firstPayPeriod = CDate("1/3/2014")

    Do Until Year(firstPayPeriod + 7) = Year(Date)
        firstPayPeriod = firstPayPeriod + 14
    Loop

'If the first PP end date this year is after the current date then we assume the most recent PP
'is pay period 26.
If firstPayPeriod >= Date Then

    RecentPP = 26
    Exit Function
    
End If

'Determine which pay period just ended.
calculatedPayPeriod = firstPayPeriod
RecentPP = 1

Do Until lastPPEndDate = calculatedPayPeriod

    calculatedPayPeriod = calculatedPayPeriod + 14
    RecentPP = RecentPP + 1
    
Loop

End Function
