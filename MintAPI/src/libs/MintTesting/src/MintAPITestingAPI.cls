VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MintAPITestingAPI"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function Tester(Optional WhatToTest As Object, Optional RegisterAsDefault As Boolean = True) As Tester
    Dim WTT As Object
    Set WTT = WhatToTest
    If WTT Is Nothing Then _
        Set WTT = globalTestingInstance
    If WTT Is Nothing Then throw Exceptions.ArgumentNullException
    If Not TypeOf WTT Is ObjectBuffer Then throw Exceptions.InvalidArgumentTypeException
    Dim t As New Tester
    Call t.Initialize(WTT)
    Set Tester = t
    If RegisterAsDefault Then _
        Set globalTestingInstance = WTT
End Function
