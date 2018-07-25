VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Tester"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim OB As ObjectBuffer

Private Sub CheckOBJ()
    
End Sub

Private Sub Class_Initialize()
    Set OB = globalTestingInstance
End Sub
Friend Sub Initialize(WhatToTest As Object)
    Set OB = WhatToTest
End Sub

Public Property Get Console() As MintAPI.Console
    
End Property
Public Property Get Files() As MintAPI.File()
    
End Property
Public Property Get Sockets() As MintAPI.Socket()
    
End Property
Public Property Get Timers() As MintAPI.Timer()
    
End Property
Public Property Get Providers() As MintAPI.IProvider()
    
End Property
Public Property Get Configuration() As MintAPI.Configuration
    
End Property

Public Sub stdout(ParamArray Args() As Variant)
    
End Sub
Public Sub stdin(ParamArray Args() As Variant)
    
End Sub
Public Sub stderr(ParamArray Args() As Variant)
    
End Sub


Public Property Get Instance() As TestingInstance
    
End Property
