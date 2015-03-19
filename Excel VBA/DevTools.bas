Attribute VB_Name = "DevTools"
Option Explicit

''
' Extensibility Library: For Meta stuff
Private Const VBA_EXTENSIBILITY_LIB As String = _
    "C:\Program Files\Common Files\Microsoft Shared\VBA\VBA6\VBE6EXT.OLB"
Private Const VBA_EXTENSIBILITY_NAME As String = "VBIDE"

''
' Scripting Runtime: For Dictionary
Private Const VBA_SCRIPTING_LIB As String = "C:\Windows\system32\scrrun.dll"
Private Const VBA_SCRIPTING_NAME As String = "Scripting"


''
' Rubberduck for testing
Private Const RUBBERDUCK_LIB As String = _
    "C:\Program Files\Rubberduck\Rubberduck\Rubberduck.tlb"
Private Const RUBBERDUCK_NAME As String = "Rubberduck"
''
' This should mimic vbext_ComponentType.  The idea is that this module
' could operate with as little setup as possible.
'
Private Enum CompenentType

    stdModule = 1
    classModule = 2
    msForm = 3
    activeXDesigner = 11
    document = 100
    
End Enum
'
' Building VBEX
' -------------
'
Public Sub BuildVBEX(ByVal sourceDir As String, ByVal buildDir As String)

    Dim buildPath As String
    buildPath = buildDir & "VBEX.xlam"

    Dim testPath As String
    testPath = buildDir & "VBEX-Testing.xlam"

    BuildAddin sourceDir & "src\", buildPath, "VBEX"
    BuildAddin sourceDir & "test\", testPath, "Testing"
    
    ' Add VBEX references
    Dim vbexWb As Workbook
    Set vbexWb = Workbooks.Open(buildPath)
    
    AddReference vbexWb.VBProject, VBA_EXTENSIBILITY_NAME, _
        FindLibVersion(VBA_EXTENSIBILITY_LIB)
    AddReference vbexWb.VBProject, VBA_SCRIPTING_NAME, VBA_SCRIPTING_LIB
    vbexWb.Save
    
    ' Add Testing References
    Dim testWb As Workbook
    Set testWb = Workbooks.Open(testPath)
    
    AddReference testWb.VBProject, "VBEX", buildPath
    AddReference testWb.VBProject, RUBBERDUCK_NAME, _
        FindLibVersion(RUBBERDUCK_LIB)
    testWb.Save
    
    ' closing testWB doesn't effect until procedure stops
    ' procedure doesn't stop until vbexWB closes
    ' vbexWB can't close until testWB closes.
    'testWb.Close savechanges:=True
    'vbexWb.Close savechanges:=True
    
End Sub
Public Sub ExportVBEX(ByVal exportDir As String)

    Dim vbexPrj As Object
    Set vbexPrj = Workbooks("VBEX.xlam").VBProject
    
    Dim srcDir As String
    srcDir = exportDir & "src\"
    
    ExportSourceFiles vbexPrj, srcDir
    
    Dim testPrj As Object
    Set testPrj = Workbooks("VBEX-Testing.xlam").VBProject
    
    Dim testDir As String
    testDir = exportDir & "test\"
    
    ExportSourceFiles testPrj, testDir

End Sub
Private Sub BuildAddin(ByVal sourceDir As String, _
        ByVal buildPath As String, ByVal projectName As String)
        
    Dim wb As Workbook
    Set wb = Workbooks.Add
    
    Dim prj As Object
    Set prj = wb.VBProject
    
    prj.Name = projectName
    
    ImportSourceFiles prj, sourceDir
    
    wb.SaveAs buildPath, FileFormat:=55
    wb.Close savechanges:=False
    
End Sub
'
' Importing VBA Files
' -------------------
'
''
' `project` is `Object` to avoid dependence
Private Sub ImportSourceFiles(ByVal project As Object, ByVal sourceDir As String)

    Dim File As String
    File = Dir(sourceDir)
    
    While (File <> "")
        project.VBComponents.Import sourceDir & File
        File = Dir
    Wend
    
End Sub
''
'
Private Function HasReference(ByVal project As Object, ByVal refName As String) As Boolean

    Dim ref As Variant
    For Each ref In project.References
    
        If ref.Name = refName Then
            HasReference = True
            Exit Function
        End If
        
    Next ref
    
    HasReference = False

End Function
''
'
Private Sub AddReference(ByVal project As Object, ByVal refName As String, _
        ByVal dllPath As String)

    If Not HasReference(project, refName) Then
        project.References.AddFromFile dllPath
    End If

End Sub
Private Function FindLibVersion(ByVal alledgedLibPath As String) As String

    Dim altLibPath As String
    altLibPath = SwitchArch(alledgedLibPath)
    
    If Dir(alledgedLibPath) <> "" Then
        FindLibVersion = alledgedLibPath
    ElseIf Dir(altLibPath) <> "" Then
        FindLibVersion = altLibPath
    Else
        ' Raise Error
    End If
    
End Function
''
' Note "Program Files (x86)" is for 32 programs if your machine is 64
' but 64 programs if your machine is 32.
Private Function SwitchArch(ByVal libPath As String) As String

    Const LOCAL_ARCH As String = "Program Files"
    Const OTHER_ARCH As String = "Program Files (x86)"
    
    If InStr(1, libPath, "(x86)") <> 0 Then
        SwitchArch = Replace$(libPath, OTHER_ARCH, LOCAL_ARCH)
    Else
        SwitchArch = Replace$(libPath, LOCAL_ARCH, OTHER_ARCH)
    End If
    
End Function

Private Function OughtExport(ByVal compType As CompenentType) As Boolean

    OughtExport = ((compType = stdModule) Or (compType = classModule))
    
End Function

Public Sub ExportSourceFiles()

    Dim destPath As String
    Dim component As VBComponent
    
    destPath = "C:\Users\danseethaler\Documents\GitHub\VBA-Git-Repo\Excel VBA\"
    
    For Each component In Application.VBE.ActiveVBProject.VBComponents
        If component.Type = vbext_ct_ClassModule Or component.Type = vbext_ct_StdModule Then
            component.Export destPath & component.Name & ToFileExtension(component.Type)
        End If
    Next
     
End Sub

Private Function ToFileExtension(vbeComponentType As vbext_ComponentType) As String
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
