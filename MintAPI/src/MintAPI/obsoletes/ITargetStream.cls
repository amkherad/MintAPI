VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IClassStream"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "ITargetStream interface is used to be inherited by streaming classes."
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"


'Ideas from Me , Almost Everywhere you can find streaming platforms!

Option Base 0
Option Explicit
Const CLASSID As String = "IClassStream"
Const CLASSALIAS As String = APPLICATIONDOMAIN & "." & CLASSID
Const CLASSVERSIONSTRING As String = "0.1"
Const CLASSALIASSTRING As String = CLASSALIAS & "/" & CLASSVERSIONSTRING


Public Enum StreamDirection
    sdInStream = 1
    sdOutStream = 2
    sdBoth = sdInStream Or sdOutStream
End Enum

Public Function GetState(ByVal Dir As StreamDirection) As Boolean: End Function

Public Sub SetOutState(Optional ByVal State As Boolean = True, Optional Reserved): End Sub
Public Sub OutStream(Data, Optional ByVal Length As Long = -1): End Sub

Public Sub SetInState(Optional ByVal State As Boolean = True, Optional Reserved): End Sub
Public Sub InStream(Data, Optional ByVal Length As Long = -1): End Sub

Public Function ReadBytes(ByVal Length As Long, Optional ByVal StartAt As Long = 0) As Byte(): End Function
Public Function Available(ByVal Direction As StreamDirection) As Long: End Function

Public Sub OpenStream(Optional ByVal Direction As StreamDirection = StreamDirection.sdBoth): End Sub
Public Sub CloseStream(Optional ByVal Direction As StreamDirection = StreamDirection.sdBoth): End Sub
