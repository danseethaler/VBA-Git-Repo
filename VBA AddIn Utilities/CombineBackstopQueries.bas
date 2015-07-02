Attribute VB_Name = "CombineBackstopQueries"
Option Explicit

Sub CombineBackstopQueries()
Dim directoryPath As String
Dim fileName As String

Application.ScreenUpdating = False

directoryPath = CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\Backstop Queries\"

fileName = Dir(directoryPath, vbReadOnly) ' + vbHidden)

'Delete empty workbooks
Do While fileName <> ""


Workbooks.Open directoryPath & fileName

If ActiveSheet.Range("B1") = " 0" Then

        ActiveWorkbook.Close
        
        Kill directoryPath & fileName
        
        Else
        
        ActiveWorkbook.Close

End If

fileName = Dir

Loop


'Merge remaining workbooks

Workbooks.Add

    ActiveWorkbook.SaveAs fileName:= _
        "C:\Users\danseethaler\Desktop\Backstop Queries\Quality Errors " & Replace(Date, "/", "-") & ".xlsx", _
        FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False

fileName = Dir(directoryPath, vbReadOnly) ' + vbHidden)

Do While fileName <> ""

If Left(fileName, 14) <> "Quality Errors" Then
Workbooks.Open directoryPath & fileName

ActiveSheet.Name = Left(ActiveWorkbook.Name, 30)

ActiveSheet.Move After:=Workbooks("Quality Errors " & Replace(Date, "/", "-") & ".xlsx").Sheets(1)

Kill directoryPath & fileName

End If

fileName = Dir

Loop

Application.DisplayAlerts = False

ActiveWorkbook.Sheets(1).Delete

Application.DisplayAlerts = True

ActiveWorkbook.Save

Application.ScreenUpdating = True

End Sub
