VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ScriptingTypes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum Scope
    sPrivate = &H0
    sPublic = &H1
    sProtected = &H2
    sFriend = &H3
    sroot = &HFF
End Enum

Public Enum ReturnType
    rtNumber = 1
    rtString = 2
    rtClass = 4
End Enum

Public Type FuncRetVal
    retType As ReturnType
    retMain As Variant
End Type
