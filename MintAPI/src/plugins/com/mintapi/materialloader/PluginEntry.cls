VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PluginEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IPlugin

Private Function IPlugin_GetObject(Name As String, Args As MintAPI.ArgumentList) As Object

End Function

Private Function IPlugin_Main(Args As MintAPI.ArgumentList) As Long
    
End Function

Private Sub IPlugin_StartupEntry(HP As MintAPI.HostProvider, StandAloneType As MintAPI.StandAloneType, Args As MintAPI.ArgumentList)
    
End Sub

Private Function IPlugin_Application() As MintAPI.IApplication
    Set IPlugin_Application = tApplication
End Function

Private Function IPlugin_GeneralProvider(ProviderEngine As String) As MintAPI.IProvider
    
End Function

Private Function IPlugin_Execute(strExecuteCommand As String, Args() As Variant) As Variant
    Select Case LCase(strExecuteCommand)
        Case "installplugin"
            
        Case "uninstallplugin"
            
        Case "registerplugin"
            
        Case ""
            
        Case ""
            
    End Select
End Function

'SupportStandAlone : False
Private Function IPlugin_SupportStandAlone() As Boolean: IPlugin_SupportStandAlone = False: End Function

