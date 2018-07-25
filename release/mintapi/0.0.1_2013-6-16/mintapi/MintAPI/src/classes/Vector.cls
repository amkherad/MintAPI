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
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"Vector"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Base 0
Option Explicit
Const CLASSID As String = "Vector"
Const CLASSALIAS As String = APPLICATIONDOMAIN & "." & CLASSID
Const CLASSVERSIONSTRING As String = "0.1"
Const CLASSALIASSTRING As String = CLASSALIAS & "/" & CLASSVERSIONSTRING


Implements ISupportSignal
Implements ISupportSlot
Implements IClassClonable
Implements IClassTexer
Implements IAliasable


Private Type PlaceHolder
    Value As Variant
End Type


Dim v() As PlaceHolder
Dim vCount As Long

Dim mySignalEmitter As New SignalEmitter


Private Sub Class_Initialize()
    Call mySignalEmitter.Initialize(Me)
End Sub

Friend Sub Initialize(Size As Long, vType As VbVarType)
    
End Sub

Public Function Vector(Size As Long, vType As VbVarType) As Vector
    
End Function

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
    
End Property
Public Property Let Item(Index As Long, Value As Variant)
    
End Property
Public Property Set Item(Index As Long, Value As Object)
    
End Property

Public Sub Swap(destinationIndex As Long, sourceIndex As Long)
    
End Sub

Public Function toList() As List
    
End Function
Public Function toCollection() As Collection
    
End Function



Private Function IClassClonable_Clone() As Object
    Set IClassClonable_Clone = Clone
End Function

Private Sub IClassTexer_fromByteArray(Value() As Byte)
    
End Sub
Private Sub IClassTexer_fromString(Value As String)
    
End Sub
Private Property Get IClassTexer_texHeader(texType As mint_TexingType) As String
    
End Property
Private Function IClassTexer_toByteArray() As Byte()
    
End Function
Private Function IClassTexer_toString() As String
    
End Function

Private Function ISupportSignal_Emitter() As SignalEmitter
    Set ISupportSignal_Emitter = mySignalEmitter
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
