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
Option Base 0
Option Explicit
Const CLASSID As String = "IPlugin"
Const CLASSALIAS As String = APPLICATIONDOMAIN & "." & CLASSID
Const CLASSVERSIONSTRING As String = "0.1"
Const CLASSALIASSTRING As String = CLASSALIAS & "/" & CLASSVERSIONSTRING


Public Enum StandAloneType
    satYes
    satNo
    satMaybe
End Enum

Public Function SupportStandAlone() As Boolean
    throw NotImplementedException
End Function
Public Sub StartupEntry(HP As HostProvider, StandAloneType As StandAloneType, Args As ArgumentList)
    throw NotImplementedException
End Sub
Public Function Main(Args As ArgumentList) As Long
    throw NotImplementedException
End Function

Public Function Application() As IApplication
    throw NotImplementedException
End Function

Public Function GeneralProvider(ProviderEngine As String) As IProvider
    throw NotImplementedException
End Function
Public Function GetObject(Name As String, Args As ArgumentList) As Object
    throw NotImplementedException
End Function
Public Function Execute(strExecuteCommand As String, Args() As Variant) As Variant
    throw NotImplementedException
End Function
