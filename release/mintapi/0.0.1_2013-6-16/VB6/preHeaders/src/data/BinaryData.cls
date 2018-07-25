VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BinaryData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@PROJECT_LICENSE
Option Explicit
Option Base 0
Implements IData
Const CLASSID As String = "BinaryData"
Private Declare Function API_VarPtrArray Lib "msvbvm60" Alias "VarPtr" (Ptr() As Any) As Long

Dim c() As Byte

Public Sub Clear()
    Erase c
End Sub
Private Sub IData_Clear(): Clear: End Sub

Private Sub IData_getCopyOfEntireDataByteArray(targetArray() As Byte)
    targetArray = c()
End Sub

Private Sub IData_getEntireDataByteArray(targetArray() As Byte)
    Call API_CopyMemory(ByVal API_VarPtrArray(targetArray()), ByVal API_VarPtrArray(c()), 4)
End Sub

Private Sub IData_ReadByteArray(targetArray() As Byte)
    
End Sub

Private Sub IData_ReadData(targetData As Variant)
    
End Sub
Private Sub IData_SetData(targetData As Variant)
    
End Sub

Private Sub IData_WriteByteArray(targetArray() As Byte)
    
End Sub

Private Property Get IData_Length() As Long
    On Error GoTo err
    IData_Length = UBound(c) - LBound(c) + 1
err:
End Property
Private Property Get IData_IsEmpty() As Boolean: IData_IsEmpty = (IData_Length <= 0): End Property
