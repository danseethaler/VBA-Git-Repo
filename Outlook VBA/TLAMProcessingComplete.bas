Attribute VB_Name = "TLAMProcessingComplete"
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
