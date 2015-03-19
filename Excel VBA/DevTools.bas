Attribute VB_Name = "DevTools"
Option Explicit

Public Sub ExportSourceFiles()

    Dim destPath As String
    Dim component As VBComponent
    Dim project As VBProject
    
    Set project = Application.VBE.VBProjects("PersonalUtilities")
    
    destPath = "C:\Users\danseethaler\Documents\GitHub\VBA-Git-Repo\Excel VBA\"
    
    For Each component In project.VBComponents
        If component.Type = vbext_ct_ClassModule Or component.Type = vbext_ct_StdModule Then
            component.Export destPath & component.Name & ToFileExtension(component.Type)
        End If
    Next
     
End Sub

Private Function ToFileExtension(vbeComponentType As vbext_ComponentType) As String

'Provide the correct file extention for the object
    Select Case vbeComponentType
        Case vbext_ComponentType.vbext_ct_ClassModule
            ToFileExtension = ".cls"
        Case vbext_ComponentType.vbext_ct_StdModule
            ToFileExtension = ".bas"
        Case vbext_ComponentType.vbext_ct_MSForm
            ToFileExtension = ".frm"
        Case vbext_ComponentType.vbext_ct_ActiveXDesigner
        Case vbext_ComponentType.vbext_ct_Document
        Case Else
            ToFileExtension = vbNullString
    End Select
     
End Function

Public Sub RefreshModulesWithGitRepo()
'Remove all modules
    Dim comp As VBComponent
    Dim project As VBProject
    Set project = Application.VBE.VBProjects("PersonalUtilities")
    
    For Each comp In project.VBComponents
        If Not comp.Name = "DevTools" And (comp.Type = vbext_ct_ClassModule Or comp.Type = vbext_ct_StdModule) Then
            project.VBComponents.Remove comp
        End If
    Next
    
'Import current files from VBA Git Repo
    Dim sourcePath As String
    Dim file As String
    
    sourcePath = "C:\Users\danseethaler\Documents\GitHub\VBA-Git-Repo\Excel VBA\"
    
    file = Dir(sourcePath)
    While (file <> vbNullString)
        project.VBComponents.Import sourcePath & file
        file = Dir
    Wend

End Sub
