VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "tApplication"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Implements IApplication

Const VERSIONSTRING As String = "0.0.1.2012"

Const COMPILEYEAR As String = "2012"
Const COMPILEDATE As String = "0/0/2012"
Const COMPILERINFO As String = "VB6 Advanced"
Const COMPILEBY As String = "Ali Mousavi Kherad | alimousavikherad@gmail.com"



Private Sub IApplication_ApplicationConfigurationChanged(ByVal E As MintAPI.EventArgs)

End Sub

Private Sub IApplication_ApplicationLanguageChanged(ByVal E As MintAPI.EventArgs)

End Sub

Private Sub IApplication_Main(args As Variant)
    'code here.
    
    'Main Application Entry...
    
    'code here.
End Sub

Private Sub IApplication_Message(ByVal E As MintAPI.ApplicationMessageEventArgs)

End Sub

Private Sub IApplication_Restart(ByVal E As MintAPI.EventArgs)

End Sub

Private Sub IApplication_Starting(ByVal E As MintAPI.EventArgs)
    'code here.
End Sub
Private Sub IApplication_Started(ByVal E As MintAPI.EventArgs)
    'code here.
End Sub
Private Sub IApplication_ApplicationInputMethodChanged(ByVal E As MintAPI.EventArgs)
    'code here.
End Sub
Private Sub IApplication_ApplicationNetworkStateChanged(ByVal E As MintAPI.EventArgs)
    'code here.
End Sub
Private Sub IApplication_Debugger(ByVal E As MintAPI.ExceptionOccuredEventArgs)
    'code here.
    Dim f As File
    Set f = File(Directory.ConcatPath(App.Path, "log.txt"))
    Call f.OpenFile(fomAppendCreate)
    If Not E.Exception Is Nothing Then
        Call f.WriteLine(E.Exception.Message)
    Else
        Call f.WriteLine("An error occured")
    End If
    Call f.CloseFile
End Sub
Private Sub IApplication_Error(ByVal E As MintAPI.ExceptionOccuredEventArgs)
    Call IApplication_Debugger(E)
    'code here.
End Sub
Private Sub IApplication_Terminating(ByVal E As MintAPI.DisposingEventArgs)
    'code here.
End Sub
Private Sub IApplication_Terminated(ByVal E As MintAPI.EventArgs)
    'code here.
End Sub

Private Function IApplication_AppDomain() As String
    IApplication_AppDomain = "com.mintapi.vb6plugin"
End Function
Private Function IApplication_ExecutionPath() As String
    IApplication_ExecutionPath = Directory.ConcatPath(App.Path, App.EXEName)
End Function
Private Function IApplication_ApplicationPath() As String
    IApplication_ApplicationPath = App.Path
End Function
Private Function IApplication_ApplicationPluginsPath() As String
    IApplication_ApplicationPluginsPath = Directory.ConcatPath(App.Path, "plugins")
End Function
Private Function IApplication_ApplicationRegistryPath(Optional ByVal System As Boolean = True) As String
    IApplication_ApplicationRegistryPath = "HKEY_LOCAL_MACHINE\Software\MintAPI\vb6plugin\"
End Function
Private Function IApplication_ApplicationTempPath() As String
    IApplication_ApplicationTempPath = Application.GetApplicationSpecifiedTempPath(False)
End Function
Private Function IApplication_Author() As String
    IApplication_Author = "Ali Mousavi Kherad"
End Function
Private Function IApplication_CompanyName() As String
    IApplication_CompanyName = "MintAPI Team"
End Function
Private Function IApplication_ApplicationLogFile() As String
    IApplication_ApplicationLogFile = Directory.ConcatPath(Application.GetAppDataPath(False), "log.txt")
End Function
Private Function IApplication_ApplicationName() As String
    IApplication_ApplicationName = "MintAPI_VB6"
End Function
Private Function IApplication_ApplicationBinPath() As String
    IApplication_ApplicationBinPath = App.Path
End Function
Private Function IApplication_ApplicationConfigurationPath() As String
    IApplication_ApplicationConfigurationPath = Directory.ConcatPath(App.Path, "config.ini")
End Function
Private Function IApplication_ApplicationConfigurationRegistryPath(Optional ByVal System As Boolean = True) As String
    IApplication_ApplicationConfigurationRegistryPath = "HKEY_LOCAL_MACHINE\Software\MintAPI\vb6plugin"
End Function
Private Function IApplication_ApplicationDataPath() As String
    IApplication_ApplicationDataPath = Application.GetAppDataPath(True)
End Function
Private Function IApplication_GUID() As String
    IApplication_GUID = "{2CA0B0AB-2115-483F-9973-8BB207060B4F}"
End Function

Private Function IApplication_HelpLink() As String
    IApplication_HelpLink = "ronote.org/help"
End Function
Private Function IApplication_LegalCopyright() As String
    IApplication_LegalCopyright = "Ali Mousavi Kherad"
End Function
Private Function IApplication_ProductCode() As String
    IApplication_ProductCode = ""
End Function
Private Function IApplication_ProductCode50chars() As String
    IApplication_ProductCode50chars = ""
End Function
Private Function IApplication_ProductName() As String
    IApplication_ProductName = App.ProductName
End Function
Private Function IApplication_ServiceWebsite() As String
    IApplication_ServiceWebsite = "support.ronote.org"
End Function
Private Function IApplication_SupportLink() As String
    IApplication_SupportLink = "support.ronote.org"
End Function
Private Function IApplication_UpdateLink() As String
    IApplication_UpdateLink = "update.ronote.org"
End Function
Private Function IApplication_VersionMajor() As Integer
    IApplication_VersionMajor = App.Major
End Function
Private Function IApplication_VersionMinor() As Long
    IApplication_VersionMinor = App.Minor
End Function
Private Function IApplication_VersionRevision() As Long
    IApplication_VersionRevision = App.Revision
End Function
Private Function IApplication_VersionString() As String
    IApplication_VersionString = App.Major & "." & App.Minor & "." & App.Revision & "." & COMPILEYEAR
End Function
Private Function IApplication_VersionTags() As String
    IApplication_VersionTags = "advanced"
End Function
Private Function IApplication_Website() As String
    IApplication_Website = "ronote.org"
End Function

