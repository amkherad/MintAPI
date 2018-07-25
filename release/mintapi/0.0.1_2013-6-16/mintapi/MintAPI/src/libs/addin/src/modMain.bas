Attribute VB_Name = "modMain"
Option Explicit

Public Const LICENSE As String = ""

Public Enum ConfigurationPlace
    cpRegistry
    cpFile
End Enum

Public C As Connect

Public Function PojectProperties(Instance As VBProject) As ProjectProperties
    Dim p As New ProjectProperties
    Call p.Initialize(Instance)
    Set PojectProperties = p
End Function

Public Function CreateLanguageFolder(Path As String) As Boolean
On Error GoTo err
    If MsgBox("Do you want to generate default application language directory instead of application path?", vbCritical Or vbYesNoCancel, "Language File") = vbYes Then
        Call Directory(Path).Create
        CreateLanguageFolder = True
    End If
err:
End Function

Public Sub Main()
    Call SpecialMethods.RegisterLibraryLicense(LICENSE, "")
    Dim tApp As New tApplication
    Call Application.StartApplication(tApp, DynamicLinkLibrary, App, Nothing)
End Sub
