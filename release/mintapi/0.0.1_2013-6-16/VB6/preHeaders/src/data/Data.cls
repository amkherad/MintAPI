VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Data"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@PROJECT_LICENSE
Option Explicit
Option Base 0
Implements IData
Const CLASSID As String = "Data"

Private Sub IData_Clear()

End Sub

Private Sub IData_getEntireDataByteArray(targetArray() As Byte)
    
End Sub

Private Sub IData_getCopyOfEntireDataByteArray(targetArray() As Byte)
    
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
    
End Property
Private Property Get IData_IsEmpty() As Boolean: IData_IsEmpty = (IData_Length <= 0): End Property
