VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Vector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"Vector"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"


'Ideas from Me , Java , Qt , .net Array Generic Class.

Option Base 0
Option Explicit
Const CLASSID As String = "Vector"
Const CLASSALIAS As String = APPLICATIONDOMAIN & "." & CLASSID
Const CLASSVERSIONSTRING As String = "0.1"
Const CLASSALIASSTRING As String = CLASSALIAS & "/" & CLASSVERSIONSTRING


Implements ISupportSignal
Implements ISupportSlot
Implements ICloneable
Implements IAliasable


Private Type PlaceHolder
    Value As Variant
End Type


Dim V() As PlaceHolder
Dim vCount As Long
Dim vType As VbVarType
Dim chkType As Boolean

Dim lockState As Boolean

Dim mySignalEmitter As New SignalEmitter

Private Sub Class_Initialize()
    Call mySignalEmitter.Initialize(Me)
    lockState = False
End Sub
Private Sub throwLockedException()
    throw Exps.InvalidStatusException("Vector is read only and locked.")
End Sub

Friend Sub Constructor0(ByVal Size As Long, Optional ByVal ArrayType As ArrayTypes = vbVariant, Optional ByVal CheckVarType As Boolean = False)
    If Size = 0 Then throw Exps.ZeroArgumentException
    If Size < 0 Then throw Exps.NegativeArgumentException
    ReDim V(Size - 1)
    vCount = Size
    
    vType = ArrayType
    chkType = CheckVarType
End Sub

Public Property Get CheckTypes() As Boolean
    
End Property
Public Property Let CheckTypes(Value As Boolean)
    
End Property

Public Function Clone() As Vector
    
End Function
Public Sub Clear()
    
End Sub

'Public Property Get Count() As Long
'    Count = vCount
'End Property
Public Property Get Length() As Long
    Length = vCount
End Property

Public Property Get Item(Index As Long) As Variant
Attribute Item.VB_UserMemId = 0
    If (Index < 0) Or (Index >= vCount) Then throw Exps.OutOfRangeException
    If IsObject(V(Index).Value) Then
        Set Item = V(Index).Value
    Else
        Item = V(Index).Value
    End If
End Property
Public Property Let Item(Index As Long, Value As Variant)
    If (Index < 0) Or (Index >= vCount) Then throw Exps.OutOfRangeException

End Property
Public Property Set Item(Index As Long, Value As Object)
    If (Index < 0) Or (Index >= vCount) Then throw Exps.OutOfRangeException
    
End Property

Public Property Get DataHandle() As Long
    
End Property
Public Property Get Handle() As Long
    
End Property

Public Sub Dispose()
    
End Sub

Public Sub Swap(destinationIndex As Long, sourceIndex As Long)
    If (destinationIndex < 0) Or (destinationIndex >= vCount) Then throw Exps.OutOfRangeException
    If (sourceIndex < 0) Or (sourceIndex >= vCount) Then throw Exps.OutOfRangeException
    
End Sub

Public Function toList() As List
    
End Function
Public Function toCollection() As Collection
    
End Function

Public Sub MakeItReadOnly()
    
End Sub
Public Function GetUnlocked() As Vector
    
End Function
    


Private Function ICloneable_Clone() As Object
    Set ICloneable_Clone = Clone
End Function

Private Function ISupportSignal_Emitter() As SignalEmitter
    Set ISupportSignal_Emitter = mySignalEmitter.ExternEmitter
End Function

Private Property Get ISupportSlot_AutoManage() As Boolean
    
End Property
Private Sub ISupportSlot_ManualManageCall(SlotName As String)
    
End Sub
Private Function ISupportSlot_Methods() As String
    
End Function

Private Property Get IAliasable_Alias() As String
    IAliasable_Alias = CLASSALIASSTRING
End Property
