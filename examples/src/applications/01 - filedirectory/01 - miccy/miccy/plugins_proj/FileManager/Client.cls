VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Client"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Const CLASSID As String = "Client"
Implements IPlugin

Private Sub IPlugin_Finilized(ByVal provider As provider)
    
End Sub
Private Sub IPlugin_Finilizing(ByVal provider As provider, canCancel As Boolean, Cancel As Boolean)
    
End Sub
Private Sub IPlugin_Initialized(ByVal provider As provider)
    
End Sub
Private Function IPlugin_provideCanUninstall() As Boolean
    
End Function
Private Function IPlugin_provideCanUpdate() As Boolean
    
End Function
Private Function IPlugin_provideHaveSettings() As Boolean
    
End Function
Private Function IPlugin_provideIcon() As stdole.Picture
    
End Function
Private Function IPlugin_provideName() As String
    IPlugin_provideName = "File Manager"
End Function
Private Function IPlugin_countPlugins() As Long
    IPlugin_countPlugins = 4
End Function
Private Function IPlugin_provideCompany() As String
    IPlugin_provideCompany = "deftro"
End Function
Private Function IPlugin_provideEmail() As String
    IPlugin_provideEmail = "pinkfluid@deftro.com"
End Function
Private Function IPlugin_provideLicence() As String
    IPlugin_provideLicence = "LGPL V3"
End Function
Private Function IPlugin_provideProgrammer() As String
    IPlugin_provideProgrammer = "Pink Fluid"
End Function
Private Function IPlugin_providePublisher() As String
    IPlugin_providePublisher = "deftro"
End Function
Private Function IPlugin_provideSettings() As Iwindow
    
End Function
Private Function IPlugin_provideToolkit() As String
    IPlugin_provideToolkit = "vb6"
End Function
Private Function IPlugin_provideUniqueID() As String
    IPlugin_provideUniqueID = "miccyfilemanager"
End Function
Private Function IPlugin_provideUpdate() As Boolean
    IPlugin_provideUpdate = False
End Function
Private Function IPlugin_provideVersion() As String
    IPlugin_provideVersion = "1.0.0.2012"
End Function
Private Function IPlugin_provideWindow(Index As Long) As String
    
End Function
Private Function IPlugin_search(searchStr As String) As Boolean
    
End Function
