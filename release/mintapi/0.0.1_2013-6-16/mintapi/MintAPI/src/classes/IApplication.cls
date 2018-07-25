VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IApplication"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Base 0
Option Explicit
Const CLASSID As String = "IApplication"
Const CLASSALIAS As String = APPLICATIONDOMAIN & "." & CLASSID
Const CLASSVERSIONSTRING As String = "0.1"
Const CLASSALIASSTRING As String = CLASSALIAS & "/" & CLASSVERSIONSTRING



Public Function ApplicationName() As String 'it's very important
    throw NotImplementedException
End Function

Public Sub Main(Args) 'VERY IMPORTANT
    throw NotImplementedException
End Sub
Public Sub Starting(ByVal E As EventArgs)
    throw NotImplementedException
End Sub
Public Sub Started(ByVal E As EventArgs)
    throw NotImplementedException
End Sub
Public Sub Restart(ByVal E As EventArgs)
    throw NotImplementedException
End Sub
Public Sub Terminating(ByVal E As DisposingEventArgs) 'VERY IMPORTANT
    throw NotImplementedException
End Sub
Public Sub Terminated(ByVal E As EventArgs)
    throw NotImplementedException
End Sub
Public Sub Error(ByVal E As ExceptionOccuredEventArgs) 'VERY IMPORTANT
    throw NotImplementedException
End Sub
Public Sub Debugger(ByVal E As ExceptionOccuredEventArgs) 'VERY IMPORTANT
    throw NotImplementedException
End Sub
Public Sub Message(ByVal E As ApplicationMessageEventArgs)
    throw NotImplementedException
End Sub

Public Function AppDomain() As String '*VERY IMPORTANT
    throw NotImplementedException
End Function

Public Function VersionString() As String 'ex : 0.0.0.1054 ultimate 'VERY IMPORTANT
    throw NotImplementedException
End Function
Public Function VersionTags() As String 'ex : ultimate plugin01 monitorplugin 'VERY IMPORTANT
    throw NotImplementedException
End Function
Public Function VersionMajor() As Integer 'VERY IMPORTANT
    throw NotImplementedException
End Function
Public Function VersionMinor() As Long 'VERY IMPORTANT
    throw NotImplementedException
End Function
Public Function VersionRevision() As Long 'VERY IMPORTANT
    throw NotImplementedException
End Function
Public Function CompanyName() As String 'it's very important 'VERY IMPORTANT
    throw NotImplementedException
End Function
Public Function ProductName() As String 'it's very important 'VERY IMPORTANT
    throw NotImplementedException
End Function
Public Function ProductCode50chars() As String 'ex : f38fdalnmmFJGSAf35097saj5oreewdalnmmFJqewcxzblkjad  'VERYIMPORTANT
    throw NotImplementedException
End Function
Public Function ProductCode() As String 'ex : f38fdalnmmFJGSAf3509 'same as 20ch of chars above
    throw NotImplementedException
End Function
Public Function Guid() As String 'ex : {5D7EA74C-BFDA-46D4-B51B-C633D678420B} 'can be empty
    throw NotImplementedException
End Function
Public Function Author() As String 'VERY IMPORTANT
    throw NotImplementedException
End Function
Public Function LegalCopyright() As String
    throw NotImplementedException
End Function
Public Function HelpLink() As String
    throw NotImplementedException
End Function
Public Function SupportLink() As String
    throw NotImplementedException
End Function
Public Function UpdateLink() As String
    throw NotImplementedException
End Function
Public Function Website() As String 'VERY IMPORTANT
    throw NotImplementedException
End Function
Public Function ServiceWebsite() As String
    throw NotImplementedException
End Function


Public Sub ApplicationNetworkStateChanged(ByVal E As EventArgs)
    throw NotImplementedException
End Sub
Public Sub ApplicationInputMethodChanged(ByVal E As EventArgs)
    throw NotImplementedException
End Sub
Public Sub ApplicationLanguageChanged(ByVal E As EventArgs)
    throw NotImplementedException
End Sub
Public Sub ApplicationConfigurationChanged(ByVal E As EventArgs)
    throw NotImplementedException
End Sub

Public Function ExecutionPath() As String 'including exefile name. 'VERY IMPORTANT
    throw NotImplementedException
End Function
Public Function ApplicationPath() As String 'excluding exefile name. 'VERY IMPORTANT
    throw NotImplementedException
End Function
Public Function ApplicationLogFile() As String 'VERY IMPORTANT
    throw NotImplementedException
End Function
Public Function ApplicationTempPath() As String 'VERY IMPORTANT
    throw NotImplementedException
End Function
Public Function ApplicationDataPath() As String 'VERY IMPORTANT
    throw NotImplementedException
End Function
Public Function ApplicationBinPath() As String 'VERY IMPORTANT
    throw NotImplementedException
End Function
Public Function ApplicationPluginsPath() As String 'VERY IMPORTANT
    throw NotImplementedException
End Function
Public Function ApplicationConfigurationPath() As String 'VERY IMPORTANT
    throw NotImplementedException
End Function

Public Function ApplicationRegistryPath(Optional ByVal System As Boolean = True) As String 'VERY IMPORTANT
    throw NotImplementedException
End Function
Public Function ApplicationConfigurationRegistryPath(Optional ByVal System As Boolean = True) As String 'VERY IMPORTANT
    throw NotImplementedException
End Function
