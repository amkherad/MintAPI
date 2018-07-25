VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IPlugin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Const CLASSID As String = "IPlugin"

Public Function provideName() As String: throw NotImplementedException: End Function
Public Function provideUniqueID() As String: throw NotImplementedException: End Function
Public Function provideCompany() As String: throw NotImplementedException: End Function
Public Function provideVersion() As String: throw NotImplementedException: End Function
Public Function provideProgrammer() As String: throw NotImplementedException: End Function
Public Function providePublisher() As String: throw NotImplementedException: End Function
Public Function provideToolkit() As String: throw NotImplementedException: End Function
Public Function countPlugins() As Long: throw NotImplementedException: End Function
Public Function provideEmail() As String: throw NotImplementedException: End Function
Public Function provideLicence() As String: throw NotImplementedException: End Function
Public Function provideWindow(Index As Long) As String: throw NotImplementedException: End Function
Public Function provideIcon() As StdPicture: throw NotImplementedException: End Function
Public Function search(searchStr As String) As Boolean: throw NotImplementedException: End Function
Public Function provideCanUninstall() As Boolean: throw NotImplementedException: End Function
Public Function provideCanUpdate() As Boolean: throw NotImplementedException: End Function
Public Function provideUpdate() As Boolean: throw NotImplementedException: End Function
Public Function provideHaveSettings() As Boolean: throw NotImplementedException: End Function
Public Function provideSettings() As Iwindow: throw NotImplementedException: End Function

Public Sub Initialized(ByVal provider As provider): throw NotImplementedException: End Sub
Public Sub Finilizing(ByVal provider As provider, canCancel As Boolean, Cancel As Boolean): throw NotImplementedException: End Sub
Public Sub Finilized(ByVal provider As provider): throw NotImplementedException: End Sub
