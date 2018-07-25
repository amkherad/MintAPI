VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClientInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Property Get ClientString() As String
    
End Property

Public Property Get IsX64Environment() As Boolean
    
End Property

Public Property Get OS() As String
    
End Property
Public Property Get BrowserName() As String
    
End Property
Public Property Get BrowserVersion() As Long
    
End Property
Public Property Get BrwserFamily() As String
    
End Property

Public Property Get Cookies() As Vector
    
End Property
