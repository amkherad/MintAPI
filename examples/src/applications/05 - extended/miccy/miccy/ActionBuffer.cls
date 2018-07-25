VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ActionBuffer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Const CLASSID As String = "ActionBuffer"

Private Type IVALUE
    Name As String
    Value As Variant
    tag As Variant
End Type

Dim v() As IVALUE
Dim vCount As Long
Dim chkexists As Boolean

Private Sub Class_Initialize()
    chkexists = False
End Sub

Public Sub CopyTo(ByVal targetActionBuffer As ActionBuffer)
    Call targetActionBuffer.Clear
    Dim i As Long
    For i = 0 To vCount - 1
        Call targetActionBuffer.append(v(i).Name, v(i).Value)
    Next
End Sub

Public Property Get CheckForExistance() As Boolean
    CheckForExistance = chkexists
End Property
Public Property Let CheckForExistance(Value As Boolean)
    chkexists = Value
End Property

Public Sub append(Name As String, Value As Variant, Optional tag As Variant = Nothing)
    If chkexists Then If getIndexOf(Name) <> -1 Then throw Exps.ItemExistsException
    ReDim Preserve v(vCount)
    v(vCount).Name = Name
    If VarType(Value) = vbObject Then
        Set v(vCount).Value = Value
    Else
            v(vCount).Value = Value
    End If
    If VarType(tag) = vbObject Then
        Set v(vCount).tag = tag
    Else
            v(vCount).tag = tag
    End If
    vCount = vCount + 1
End Sub
Public Sub Clear()
    vCount = 0
    Erase v()
End Sub

Public Property Get Count() As Long
    Count = vCount
End Property

Public Function getIndexOf(Name As String) As Long
    Dim i As Long
    For i = 0 To vCount - 1
        If v(i).Name = Name Then
            getIndexOf = i
            Exit Function
        End If
    Next
    getIndexOf = -1
End Function

Public Property Get ItemName(Index As Long) As String
    ItemName = v(Index).Name
End Property
Public Property Let ItemName(Index As Long, Value As String)
    v(Index).Name = Value
End Property

Public Function getVarType(Index As Long) As VbVarType
    getVarType = VarType(v(Index).Value)
End Function
Public Function getTagVarType(Index As Long) As VbVarType
    getTagVarType = VarType(v(Index).tag)
End Function
Public Property Get Item(Index As Long) As Variant
Attribute Item.VB_UserMemId = 0
    If VarType(v(Index).Value) = vbObject Then
        Set Item = v(Index).Value
    Else
            Item = v(Index).Value
    End If
End Property
Public Property Let Item(Index As Long, Value As Variant)
    If VarType(Value) = vbObject Then
        Set v(Index).Value = Value
    Else
            v(Index).Value = Value
    End If
End Property
Public Property Set Item(Index As Long, Value As Variant)
    Set v(Index).Value = Value
End Property

Public Property Get ItemTag(Index As Long) As Variant
    If VarType(v(Index).Value) = vbObject Then
        Set ItemTag = v(Index).tag
    Else
            ItemTag = v(Index).tag
    End If
End Property
Public Property Let ItemTag(Index As Long, Value As Variant)
    If VarType(Value) = vbObject Then
        Set v(Index).tag = Value
    Else
            v(Index).tag = Value
    End If
End Property
Public Property Set ItemTag(Index As Long, Value As Variant)
    Set v(Index).tag = Value
End Property

Public Property Get isTagged(Index As Long)
    If VarType(v(Index).tag) = vbObject Then
        isTagged = Not v(Index).tag Is Nothing
    End If
    isTagged = True
End Property

Public Property Get ItemByName(Name As String) As Variant
    Dim Index As Long
    Index = getIndexOf(Name)
    If Index = -1 Then throw Exps.ItemNotExistsException
    If VarType(v(Index).Value) = vbObject Then
        Set ItemByName = v(Index).Value
    Else
            ItemByName = v(Index).Value
    End If
End Property
Public Property Let ItemByName(Name As String, Value As Variant)
    Dim Index As Long
    Index = getIndexOf(Name)
    If Index = -1 Then throw Exps.ItemNotExistsException
    If VarType(Value) = vbObject Then
        Set v(Index).Value = Value
    Else
            v(Index).Value = Value
    End If
End Property
Public Property Set ItemByName(Name As String, Value As Variant)
    Dim Index As Long
    Index = getIndexOf(Name)
    If Index = -1 Then throw Exps.ItemNotExistsException
    Set v(Index).Value = Value
End Property
