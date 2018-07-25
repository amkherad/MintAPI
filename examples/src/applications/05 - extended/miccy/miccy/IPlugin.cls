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
Option Explicit
Const CLASSID As String = "IPlugin"

Public Function provideName() As String: throw Exps.NotImplementedException: End Function
Public Function provideUniqueID() As String: throw Exps.NotImplementedException: End Function
Public Function provideCompany() As String: throw Exps.NotImplementedException: End Function
Public Function provideVersion() As String: throw Exps.NotImplementedException: End Function
Public Function provideProgrammer() As String: throw Exps.NotImplementedException: End Function
Public Function providePublisher() As String: throw Exps.NotImplementedException: End Function
Public Function provideToolkit() As String: throw Exps.NotImplementedException: End Function
Public Function countPlugins() As Long: throw Exps.NotImplementedException: End Function
Public Function provideEmail() As String: throw Exps.NotImplementedException: End Function
Public Function provideLicence() As String: throw Exps.NotImplementedException: End Function
Public Function provideWindow(Index As Long) As String: throw Exps.NotImplementedException: End Function
