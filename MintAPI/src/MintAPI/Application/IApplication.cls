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
Attribute VB_Ext_KEY = "Member0" ,"Guid"
Attribute VB_Ext_KEY = "Member1" ,"Version"
Attribute VB_Ext_KEY = "Member2" ,"Debugger"
Option Base 0
Option Explicit
Const CLASSID As String = "IApplication"
Const CLASSALIAS As String = APPLICATIONDOMAIN & "." & CLASSID
Const CLASSVERSIONSTRING As String = "0.1"
Const CLASSALIASSTRING As String = CLASSALIAS & "/" & CLASSVERSIONSTRING

Public Type ApplicationInfo
    Name As String
    Alias As String
    Version As Version
    CompanyName As String
    ProductName As String
    LegalCopyright As String
    ProductCode50chars As String
    ProductCode As String
    Guid As Guid
    
    ConfigurationPath As String
    
    Website As String
    HelpURL As String
    SupportURL As String
    UpdateURL As String
    DocumentationURL As String
    ServiceURL As String
End Type

Public Function GetApplicationInfo() As ApplicationInfo: End Function

Public Sub Main(ByRef Arguments As Variant): End Sub
Public Sub Starting(ByVal E As EventArgs): End Sub
Public Sub Started(ByVal E As EventArgs): End Sub
Public Sub Restart(ByVal E As EventArgs): End Sub
Public Sub Terminating(ByVal E As CancelEventArgs): End Sub
Public Sub Terminated(ByVal E As EventArgs): End Sub
Public Sub Error(ByVal E As ExceptionOccuredEventArgs): End Sub
Public Sub Debugger(ByVal E As DebuggerNotificationEventArgs): End Sub

Public Sub EventMessage(ByVal E As ApplicationMessageEventArgs): End Sub
'Public Sub ApplicationNetworkStateChanged(ByVal E As EventArgs): End Sub
'Public Sub ApplicationInputMethodChanged(ByVal E As EventArgs): End Sub
'Public Sub ApplicationLanguageChanged(ByVal E As EventArgs): End Sub
'Public Sub ApplicationConfigurationChanged(ByVal E As EventArgs): End Sub
