Attribute VB_Name = "NotificationEmails"
Option Explicit

Sub TLAMProcessingComplete()
    Dim Outlook As Application
    Dim myItem As MailItem
    Dim PP As Integer
    
    Set Outlook = CreateObject("Outlook.Application")
    Set myItem = Outlook.CreateItemFromTemplate("\\CHQPVUN0066\FINUSR\SHARED\FIN_PYRL\2_Payroll Time & Labor Absence Management\Desk Manual (Information)\Email Templates\TL and AM Processing Complete.oft")
    
    myItem.HTMLBody = Replace(myItem.HTMLBody, "PP20", "PP" & RecentPP())
    myItem.Display
    
End Sub

Sub MondayHolidayEmailReminder()
    Dim Outlook As Application
    Dim myItem As MailItem
    Dim PP As Integer
    
    Set Outlook = GetObject(, "Outlook.Application")
    Set myItem = Outlook.CreateItemFromTemplate("\\CHQPVUN0066\FINUSR\SHARED\FIN_PYRL\2_Payroll Time & Labor Absence Management\Desk Manual (Information)\Time America Processing Files\Early Payroll Deadline Due to Monday Holiday.oft")
    
    myItem.HTMLBody = Replace(myItem.HTMLBody, "PP", CurrentPP())
    myItem.HTMLBody = Replace(myItem.HTMLBody, "HolidayDate", Format(HolidayDate(), "dddd, mmm d, yyyy"))
    myItem.HTMLBody = Replace(myItem.HTMLBody, "SaturdayDueDate", Format(SaturdayDueDate(), "dddd, mmm d, yyyy"))
    myItem.HTMLBody = Replace(myItem.HTMLBody, "PayPeriodEndDate", Format(PayPeriodEndDate(), "dddd, mmm d, yyyy"))
    myItem.Display
    
End Sub

Sub OffCycleCheckReport()
    Dim Outlook As Application
    Dim myItem As MailItem
    
    Set Outlook = CreateObject("Outlook.Application")
    Set myItem = Outlook.CreateItemFromTemplate("\\CHQPVUN0066\FINUSR\SHARED\FIN_PYRL\2_Payroll Time & Labor Absence Management\Desk Manual (Information)\Email Templates\Off-Cycle Check Report - 12 Month Rolling.oft")
    
    myItem.HTMLBody = Replace(myItem.HTMLBody, "currentmonth", InputBox("What month has been added to this report?"))
    myItem.Attachments.Add "C:\Users\danseethaler\Dropbox\Work\Current Projects\Dashboards\Off-Cycle Dashboard\Off-Cycle Check Dashboard - 12 Month Rolling.xlsm"
    myItem.Display
    
    MsgBox ("Make sure to add this report to the SharePoint Team site.")
    
End Sub

Function CurrentPP() As Integer
Dim PP01 As Date
Dim lastPPEndDate As Date
Dim currentPPEndDate As Date

'Get the date of the last day of last pay period
lastPPEndDate = Date - (Date - CDate("8/15/2014")) Mod 14

'Check if today is the last day of the pay period.
If lastPPEndDate <> Date Then
        currentPPEndDate = lastPPEndDate + 14
    Else
        currentPPEndDate = lastPPEndDate
End If

'Determine the PPEnd date for PP01 of the current year.
PP01 = CDate("1/3/2014")
Do Until Year(PP01 + 7) = Year(Date)
    PP01 = PP01 + 14
Loop

'Determine the pay period number of the current pay period.
CurrentPP = 1
Do Until currentPPEndDate = PP01
    PP01 = PP01 + 14
    CurrentPP = CurrentPP + 1
Loop

End Function

Function SaturdayDueDate() As Date
Dim PP01 As Date
Dim lastPPEndDate As Date
Dim currentPPEndDate As Date

'Get the date of the last day of last pay period
lastPPEndDate = Date - (Date - CDate("8/15/2014")) Mod 14

'Check if today is the last day of the pay period.
If lastPPEndDate <> Date Then
        currentPPEndDate = lastPPEndDate + 14
    Else
        currentPPEndDate = lastPPEndDate
End If

SaturdayDueDate = currentPPEndDate + 1

End Function

Function HolidayDate() As Date
Dim PP01 As Date
Dim lastPPEndDate As Date
Dim currentPPEndDate As Date

'Get the date of the last day of last pay period
lastPPEndDate = Date - (Date - CDate("8/15/2014")) Mod 14

'Check if today is the last day of the pay period.
If lastPPEndDate <> Date Then
        currentPPEndDate = lastPPEndDate + 14
    Else
        currentPPEndDate = lastPPEndDate
End If

HolidayDate = currentPPEndDate + 3

End Function

Function PayPeriodEndDate() As Date
Dim PP01 As Date
Dim lastPPEndDate As Date
Dim currentPPEndDate As Date

'Get the date of the last day of last pay period
lastPPEndDate = Date - (Date - CDate("8/15/2014")) Mod 14

'Check if today is the last day of the pay period.
If lastPPEndDate <> Date Then
        PayPeriodEndDate = lastPPEndDate + 14
    Else
        PayPeriodEndDate = lastPPEndDate
End If

End Function

