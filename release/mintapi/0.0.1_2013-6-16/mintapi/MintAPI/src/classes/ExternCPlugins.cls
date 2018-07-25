VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ExternCPlugins"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Base 0
Option Explicit
Const CLASSID As String = "ExternCPlugins"
Const CLASSALIAS As String = APPLICATIONDOMAIN & "." & CLASSID
Const CLASSVERSIONSTRING As String = "0.1"
Const CLASSALIASSTRING As String = CLASSALIAS & "/" & CLASSVERSIONSTRING


Implements IPlugin
Implements IAliasable


Dim tApp As New ExternCPlugins_Application

Dim mPath As String
Dim mhndl As Long

Public Sub Initialize(Path As String)

End Sub

Private Function IPlugin_Application() As IApplication
    
End Function

Private Function IPlugin_Command(Arg As Variant) As Variant
    
End Function

Private Function IPlugin_Execute(strExecuteCommand As String, Args() As Variant) As Variant
    
End Function

Private Function IPlugin_GeneralProvider(ProviderEngine As String) As IProvider
    
End Function

Private Function IPlugin_GetObject(Name As String, Args As ArgumentList) As Object
    
End Function

Private Function IPlugin_Main(Args As ArgumentList) As Long
    
End Function

Private Sub IPlugin_StartupEntry(HP As HostProvider, StandAloneType As StandAloneType, Args As ArgumentList)
    
End Sub

Private Function IPlugin_SupportStandAlone() As Boolean
    
End Function


Private Property Get IAliasable_Alias() As String
    IAliasable_Alias = CLASSALIASSTRING
End Property
