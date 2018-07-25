VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ProjectProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim Instance As VBIDE.VBProject

Friend Sub Initialize(VBInstance As VBIDE.VBProject)
    Set Instance = VBInstance
End Sub

Public Property Get ConfigurationPlace() As ConfigurationPlace
    If Instance Is Nothing Then Exit Property
    ConfigurationPlace = IIf(Instance.ReadProperty("MintAPI_DesignTools", "ConfigurationPlace") = "registry", cpRegistry, cpFile)
End Property
Public Property Let ConfigurationPlace(Value As ConfigurationPlace)
    If Instance Is Nothing Then Exit Property
    Call Instance.WriteProperty("MintAPI_DesignTools", "ConfigurationPlace", IIf(Value = cpRegistry, "registry", "file"))
End Property

Public Property Get GenerateConfigurationFile() As Boolean
    If Instance Is Nothing Then Exit Property
    ConfigurationPlace = Instance.ReadProperty("MintAPI_DesignTools", "GenerateConfigurationFile") = "true"
End Property
Public Property Let GenerateConfigurationFile(Value As Boolean)
    If Instance Is Nothing Then Exit Property
    Call Instance.WriteProperty("MintAPI_DesignTools", "GenerateConfigurationFile", IIf(Value, "true", "false"))
End Property

Public Property Get LanguageDirectory() As String
    If Instance Is Nothing Then Exit Property
    LanguageDirectory = Instance.ReadProperty("MintAPI_DesignTools", "LanguageDirectory")
End Property
Public Property Let LanguageDirectory(Value As String)
    If Instance Is Nothing Then Exit Property
    Call Instance.WriteProperty("MintAPI_DesignTools", "LanguageDirectory", Value)
End Property

Public Function GetDirectory(dir As String) As String
On Error GoTo err
    If Instance Is Nothing Then GetDirectory = dir: Exit Function
    If Not Instance.Saved Then GetDirectory = dir:  Exit Function
    Dim s As String
    s = File(Instance.FileName).Directory.Path
    If s = "" Then s = File(Instance.BuildFileName).Directory.Path
    If s = "" Then Exit Function
    GetDirectory = Replace(dir, "[STARTUPPATH]", s)
err:
End Function
