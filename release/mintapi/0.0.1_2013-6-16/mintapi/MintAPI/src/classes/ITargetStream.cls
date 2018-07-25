VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ITargetStream"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "ITargetStream interface is used to be inherited by streaming classes."
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Base 0
Option Explicit
Const CLASSID As String = "ITargetStream"
Const CLASSALIAS As String = APPLICATIONDOMAIN & "." & CLASSID
Const CLASSVERSIONSTRING As String = "0.1"
Const CLASSALIASSTRING As String = CLASSALIAS & "/" & CLASSVERSIONSTRING


Public Enum StreamDirection
    sdInStream = 1
    sdOutStream = 2
    sdBoth = sdInStream Or sdOutStream
End Enum

Public Function getState(Dir As StreamDirection) As Boolean
    throw NotImplementedException
End Function
Public Sub setOutState(Optional ByVal State As Boolean = True, Optional Reserved)
    throw NotImplementedException
End Sub
Public Sub outStream(Data, Optional Length As Long = -1)
    throw NotImplementedException
End Sub
Public Sub setInState(Optional ByVal State As Boolean = True, Optional Reserved)
    throw NotImplementedException
End Sub
Public Function readBytes(Length As Long, Optional StartAt As Long = 0) As Byte()
    throw NotImplementedException
End Function
Public Sub inStream(Data, Optional Length As Long = -1)
    throw NotImplementedException
End Sub
Public Function Available(direction As StreamDirection) As Long
    throw NotImplementedException
End Function

Public Sub OpenStream(Optional direction As StreamDirection = StreamDirection.sdBoth)
    throw NotImplementedException
End Sub
Public Sub CloseStream(Optional direction As StreamDirection = StreamDirection.sdBoth)
    throw NotImplementedException
End Sub
