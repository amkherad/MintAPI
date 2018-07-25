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
Attribute VB_Description = "IPlugin class is used to be inherited by base plugin class locally called [PluginInstance] for applications to support loading plugins."
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"Application"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"


'Ideas from Me , Qt

Option Base 0
Option Explicit
Const CLASSID As String = "IPlugin"
Const CLASSALIAS As String = APPLICATIONDOMAIN & "." & CLASSID
Const CLASSVERSIONSTRING As String = "0.1"
Const CLASSALIASSTRING As String = CLASSALIAS & "/" & CLASSVERSIONSTRING



'Public Sub StartupEntry(ByVal Provider As Provider, ByVal StandAloneType As StandAloneType, ByVal Args As ArgumentList): End Sub
'Public Function Main(ByVal Args As ArgumentList) As Long: End Function
'
'Public Function GeneralProvider(ByVal ProviderEngine As String) As IProvider: End Function
'Public Function GetObject(ByVal Name As String, ByVal Args As ArgumentList) As Object: End Function
'Public Function Execute(ByVal strExecuteCommand As String, Args() As Variant) As Variant: End Function
