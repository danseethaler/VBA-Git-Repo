Attribute VB_Name = "NotificationEmails"
Option Explicit

Sub TLAMProcessingComplete()
    Dim myolApp As Application
    Dim myItem As MailItem
    Dim PP As Integer
    
    Set myolApp = CreateObject("Outlook.Application")
    Set myItem = myolApp.CreateItemFromTemplate("\\CHQPVUN0066\FINUSR\SHARED\FIN_PYRL\2_Payroll Time & Labor Absence Management\Desk Manual (Information)\TL and AM Processing Complete.oft")
    
    PP = InputBox("Which PP is this for?")
    
    myItem.HTMLBody = Replace(myItem.HTMLBody, "PP20", "PP0" & PP)
    myItem.Display
    
End Sub

Sub OffCycleCheckReport()
    Dim myolApp As Application
    Dim myItem As MailItem
    
    Set myolApp = CreateObject("Outlook.Application")
    Set myItem = myolApp.CreateItemFromTemplate("\\CHQPVUN0066\FINUSR\SHARED\FIN_PYRL\2_Payroll Time & Labor Absence Management\Desk Manual (Information)\Off-Cycle Check Report - 12 Month Rolling.oft")
    
    myItem.HTMLBody = Replace(myItem.HTMLBody, "currentmonth", InputBox("What month has been added to this report?"))
    myItem.Attachments.Add "C:\Users\danseethaler\Dropbox\Work\Current Projects\Dashboards\Off-Cycle Dashboard\Off-Cycle Check Dashboard - 12 Month Rolling.xlsm"
    myItem.Display
    
    MsgBox ("Go ahead. Add some commentary.")
    
End Sub

Sub NewEmailToRecipients()
    Dim currentEmail As MailItem
    Dim outlookExplorer As Outlook.Explorer
    Dim selection As Outlook.selection
    Dim newEmail As MailItem
    
    Set outlookExplorer = Outlook.ActiveExplorer
    Set selection = outlookExplorer.selection
    
    Set currentEmail = selection.Item(1)
    
    Debug.Print currentEmail.Sender
    
    Set newEmail = CreateItem(olMailItem)
    
    If currentEmail.Sender <> "Dan Seethaler" Then
        newEmail.To = Replace(currentEmail.To, "Dan Seethaler", currentEmail.Sender)
    Else
        newEmail.To = currentEmail.To
    End If
    
    newEmail.CC = currentEmail.CC
    newEmail.Display

End Sub
