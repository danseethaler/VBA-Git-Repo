Attribute VB_Name = "CombineBackstopQueries"
Option Explicit

Sub CombineBackstopQueries()
Dim DirectoryPath As String
Dim FileName As String

Application.ScreenUpdating = False

DirectoryPath = CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\Backstop Queries\"

FileName = Dir(DirectoryPath, vbReadOnly) ' + vbHidden)

'Delete empty workbooks
Do While FileName <> ""


Workbooks.Open DirectoryPath & FileName

If ActiveSheet.Range("B1") = " 0" Then

        ActiveWorkbook.Close
        
        Kill DirectoryPath & FileName
        
        Else
        
        ActiveWorkbook.Close

End If

FileName = Dir

Loop


'Merge remaining workbooks

Workbooks.Add

    ActiveWorkbook.SaveAs FileName:= _
        "C:\Users\danseethaler\Desktop\Backstop Queries\Quality Errors " & Replace(Date, "/", "-") & ".xlsx", _
        FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False

FileName = Dir(DirectoryPath, vbReadOnly) ' + vbHidden)

Do While FileName <> ""

If Left(FileName, 14) <> "Quality Errors" Then
Workbooks.Open DirectoryPath & FileName

ActiveSheet.Name = Left(ActiveWorkbook.Name, 30)

ActiveSheet.Move After:=Workbooks("Quality Errors " & Replace(Date, "/", "-") & ".xlsx").Sheets(1)

Kill DirectoryPath & FileName

End If

FileName = Dir

Loop

Application.DisplayAlerts = False

ActiveWorkbook.Sheets(1).Delete

Application.DisplayAlerts = True

ActiveWorkbook.Save

Application.ScreenUpdating = True

End Sub
