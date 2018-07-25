VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ExternCPlugins_Application"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Base 0
Option Explicit
Const CLASSID As String = "ExternCPlugins_Application"
Const CLASSALIAS As String = APPLICATIONDOMAIN & "." & CLASSID
Const CLASSVERSIONSTRING As String = "0.1"
Const CLASSALIASSTRING As String = CLASSALIAS & "/" & CLASSVERSIONSTRING


Implements IApplication
Implements IAliasable


Dim hndl As Long

Friend Sub Initialize(ModuleHandle As Long)
    hndl = ModuleHandle
End Sub

Private Function IApplication_AppDomain() As String

End Function

Private Function IApplication_ApplicationBinPath() As String

End Function

Private Sub IApplication_ApplicationConfigurationChanged(ByVal E As EventArgs)

End Sub

Private Function IApplication_ApplicationConfigurationPath() As String

End Function

Private Function IApplication_ApplicationConfigurationRegistryPath(Optional ByVal System As Boolean = True) As String

End Function

Private Function IApplication_ApplicationDataPath() As String

End Function

Private Sub IApplication_ApplicationInputMethodChanged(ByVal E As EventArgs)

End Sub

Private Sub IApplication_ApplicationLanguageChanged(ByVal E As EventArgs)

End Sub

Private Function IApplication_ApplicationLogFile() As String

End Function

Private Function IApplication_ApplicationName() As String

End Function

Private Sub IApplication_ApplicationNetworkStateChanged(ByVal E As EventArgs)

End Sub

Private Function IApplication_ApplicationPath() As String

End Function

Private Function IApplication_ApplicationPluginsPath() As String

End Function

Private Function IApplication_ApplicationRegistryPath(Optional ByVal System As Boolean = True) As String

End Function

Private Function IApplication_ApplicationTempPath() As String

End Function

Private Function IApplication_Author() As String

End Function

Private Function IApplication_CompanyName() As String

End Function

Private Sub IApplication_Debugger(ByVal E As ExceptionOccuredEventArgs)

End Sub

Private Sub IApplication_Error(ByVal E As ExceptionOccuredEventArgs)

End Sub

Private Function IApplication_ExecutionPath() As String

End Function

Private Function IApplication_GUID() As String

End Function

Private Function IApplication_HelpLink() As String

End Function

Private Function IApplication_LegalCopyright() As String

End Function

Private Sub IApplication_Main(Args As Variant)

End Sub

Private Sub IApplication_Message(ByVal E As ApplicationMessageEventArgs)
    
End Sub

Private Function IApplication_ProductCode() As String

End Function

Private Function IApplication_ProductCode50chars() As String

End Function

Private Function IApplication_ProductName() As String

End Function

Private Sub IApplication_Restart(ByVal E As EventArgs)

End Sub

Private Function IApplication_ServiceWebsite() As String

End Function

Private Sub IApplication_Started(ByVal E As EventArgs)

End Sub

Private Sub IApplication_Starting(ByVal E As EventArgs)

End Sub

Private Function IApplication_SupportLink() As String

End Function

Private Sub IApplication_Terminated(ByVal E As EventArgs)

End Sub

Private Sub IApplication_Terminating(ByVal E As DisposingEventArgs)

End Sub

Private Function IApplication_UpdateLink() As String

End Function

Private Function IApplication_VersionMajor() As Integer

End Function

Private Function IApplication_VersionMinor() As Long

End Function

Private Function IApplication_VersionRevision() As Long

End Function

Private Function IApplication_VersionString() As String

End Function

Private Function IApplication_VersionTags() As String

End Function

Private Function IApplication_Website() As String

End Function


Private Property Get IAliasable_Alias() As String
    IAliasable_Alias = CLASSALIASSTRING
End Property
