VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ArgumentList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'@PROJECT_LICENSE

Option Base 0
Option Explicit
Const CLASSID As String = "ArgumentList"
Const CLASSALIAS As String = APPLICATIONDOMAIN & "." & CLASSID
Const CLASSVERSIONSTRING As String = "0.1"
Const CLASSALIASSTRING As String = CLASSALIAS & "/" & CLASSVERSIONSTRING


Const SIGNAL_APPEND As String = "append"
Const SIGNAL_REMOVE As String = "remove"
Const SIGNAL_CLEAR As String = "clear"

Const SLOT_CLEAR As String = "clear"

Implements IClassClonable
Implements ISupportSignal
Implements ISupportSlot
Implements IAliasable

Private Type ArgumentListLST
    Name As String
    Value As Variant
End Type

Dim Args() As ArgumentListLST
Dim ArgsCount As Long

Dim Locked As Long
Dim Keys() As String
Dim KeysCount As Long

Dim mySignalEmitter As New SignalEmitter

Private Sub Class_Initialize()
    Call mySignalEmitter.Initialize(Me)
End Sub

Friend Sub InitializeW(Args())
    Dim i As Long
    For i = 0 To ArraySize(Args) - 1
        Call AppendArgument(Args(i))
    Next
End Sub
Friend Sub Initialize(Args() As Argument)
    Dim i As Long
    For i = 0 To ArraySize(Args) - 1
        Call AppendArgument(Args(i).Value, Args(i).Name)
    Next
End Sub

Friend Sub LockArgs(Key As String)
    ReDim Preserve Keys(KeysCount)
    Keys(KeysCount) = Key
    KeysCount = KeysCount + 1
    Locked = Locked + 1
End Sub
Friend Sub UnlockArgs(Key As String)
    If KeysCount > 0 Then
        If Keys(KeysCount - 1) = Key Then
            KeysCount = KeysCount - 1
            ReDim Preserve Keys(KeysCount - 1)
            Locked = Locked - 1
        End If
    End If
End Sub

Public Sub AppendArgument(Arg, Optional Name As String)
    If Locked > 1 Then throw InvalidStatusException
    If IndexOf(Name) >= 0 Then throw ItemExistsException("Argument with the same name already exists.")
    ReDim Preserve Args(ArgsCount)
    If VarType(Arg) = VBObject Then
        Set Args(ArgsCount).Value = Arg
    Else
            Args(ArgsCount).Value = Arg
    End If
    Args(ArgsCount).Name = Name
    ArgsCount = ArgsCount + 1
    Call mySignalEmitter.Emit(SIGNAL_APPEND)
End Sub

Public Function RemoveArgumentbyIndex(Index As Long) As Boolean
    If Index < 0 Or Index >= ArgsCount Then throw OutOfRangeException
    Dim i As Long
    For i = Index To ArgsCount - 2
        Args(i) = Args(i + 1)
    Next
    ArgsCount = ArgsCount - 1
    ReDim Preserve Args(ArgsCount - 1)
    Call mySignalEmitter.Emit(SIGNAL_REMOVE)
End Function
Public Function RemoveArgument(Arg) As Boolean
    If Locked > 1 Then throw InvalidStatusException
    Dim Index As Long, vType As VbVarType
    If VBA.IsObject(Arg) Then
        For Index = 0 To ArgsCount - 1
            If VBA.IsObject(Args(Index).Value) Then
                If Arg Is Args(Index).Value Then
                    RemoveArgument = RemoveArgumentbyIndex(Index)
                    Exit Function
                End If
            End If
        Next
    Else
        For Index = 0 To ArgsCount - 1
            vType = VarType(Args(Index).Value)
            If Not (vType = VBObject Or vType = vbArray) Then
                If Arg = Args(Index).Value Then
                    RemoveArgument = RemoveArgumentbyIndex(Index)
                    Exit Function
                End If
            End If
        Next
    End If
End Function
Public Function RemoveArgumentbyName(Name As String) As Boolean
    If Locked > 1 Then throw InvalidStatusException
    Dim Index As Long
    Index = IndexOf(Name)
    If Index = -1 Then throw ItemNotExistsException
    
    RemoveArgumentbyName = RemoveArgumentbyIndex(Index)
End Function


Public Function NameOf(Index As Long) As String
    If Index < 0 Or Index >= ArgsCount Then throw OutOfRangeException
    NameOf = Args(Index).Name
End Function
Public Sub SetNameOf(Index As Long, Name As String)
    If Index < 0 Or Index >= ArgsCount Then throw OutOfRangeException
    Args(Index).Name = Name
End Sub
Public Function TypeName(Name As String) As String
    Dim Index As Long
    Index = IndexOf(Name)
    If Index = -1 Then throw ItemNotExistsException
    TypeName = VBA.TypeName(Args(Index).Value)
End Function
Public Function TypeNamebyIndex(Index As Long) As String
    If Index < 0 Or Index >= ArgsCount Then throw OutOfRangeException
    TypeNamebyIndex = VBA.TypeName(Args(Index).Value)
End Function

Public Function IndexOf(Name As String) As Long
    Dim i As Long
    For i = 0 To ArgsCount - 1
        If Args(i).Name = Name Then
            IndexOf = i
            Exit Function
        End If
    Next
    IndexOf = -1
End Function
Public Function Exists(Name As String) As Boolean
    Exists = (IndexOf(Name) <> -1)
End Function

Public Property Get Count() As Long
    Count = ArgsCount
End Property

Public Sub Clear()
    If Locked > 0 Then throw InvalidStatusException
    Erase Args
    ArgsCount = 0
    Call mySignalEmitter.Emit(SIGNAL_CLEAR)
End Sub

Public Function ArgumentType(IndexKey) As VbVarType
    Dim Index As Long
    Select Case VarType(IndexKey)
        Case vbLong, vbInteger, vbDouble, vbSingle, vbByte
            Index = CLng(IndexKey)
        Case Else
            Index = IndexOf(CStr(IndexKey))
    End Select
    If Index < 0 Or Index >= ArgsCount Then throw ItemNotExistsException
    ArgumentType = VarType(Args(Index).Value)
End Function
Public Function IsObject(IndexKey) As VbVarType
    Dim Index As Long
    Select Case VarType(IndexKey)
        Case vbLong, vbInteger, vbDouble, vbSingle, vbByte
            Index = CLng(IndexKey)
        Case Else
            Index = IndexOf(CStr(IndexKey))
    End Select
    If Index < 0 Or Index >= ArgsCount Then throw OutOfRangeException
    IsObject = VBA.IsObject(Args(Index).Value)
End Function
Public Function IsArray(IndexKey) As VbVarType
    Dim Index As Long
    Select Case VarType(IndexKey)
        Case vbLong, vbInteger, vbDouble, vbSingle, vbByte
            Index = CLng(IndexKey)
        Case Else
            Index = IndexOf(CStr(IndexKey))
    End Select
    If Index < 0 Or Index >= ArgsCount Then throw OutOfRangeException
    IsArray = VBA.IsArray(Args(Index).Value)
End Function

Public Property Get Item(IndexKey) As Variant
Attribute Item.VB_UserMemId = 0
    If Locked > 1 Then throw InvalidStatusException
    Dim Index As Long
    Select Case VarType(IndexKey)
        Case vbLong, vbInteger, vbDouble, vbSingle, vbByte
            Index = CLng(IndexKey)
        Case Else
            Index = IndexOf(CStr(IndexKey))
    End Select
    If Index < 0 Or Index >= ArgsCount Then throw ItemNotExistsException
    If VBA.IsObject(Args(Index).Value) Then
        Set Item = Args(Index).Value
    Else
            Item = Args(Index).Value
    End If
End Property
Public Property Let Item(IndexKey, Value As Variant)
    If Locked > 0 Then throw InvalidStatusException
    Dim Index As Long
    Select Case VarType(IndexKey)
        Case vbLong, vbInteger, vbDouble, vbSingle, vbByte
            Index = CLng(IndexKey)
        Case Else
            Index = IndexOf(CStr(IndexKey))
            If Index = -1 Then
                Call AppendArgument(Value, CStr(IndexKey))
                Exit Property
            End If
    End Select
    If Index < 0 Or Index >= ArgsCount Then throw ItemNotExistsException
    If VBA.IsObject(Value) Then
        Set Args(Index).Value = Value
    Else
            Args(Index).Value = Value
    End If
End Property
Public Property Set Item(IndexKey, Value As Object)
    If Locked > 0 Then throw InvalidStatusException
    Dim Index As Long
    Select Case VarType(IndexKey)
        Case vbLong, vbInteger, vbDouble, vbSingle, vbByte
            Index = CLng(IndexKey)
        Case Else
            Index = IndexOf(CStr(IndexKey))
            If Index = -1 Then
                Call AppendArgument(Value, CStr(IndexKey))
                Exit Property
            End If
    End Select
    If Index < 0 Or Index >= ArgsCount Then throw ItemNotExistsException
    Set Args(Index).Value = Value
End Property

Public Function thrownot(HowMany As Long)
    If ArgsCount > HowMany Then throw TooManyArgumentsException
    If ArgsCount < HowMany Then throw NotEnoughArgumentsException
End Function

Public Function Clone() As ArgumentList
    Dim i As Long, retAL As New ArgumentList
    For i = 0 To ArgsCount - 1
        Call retAL.AppendArgument(Args(i).Value, Args(i).Name)
    Next
End Function

Private Function IClassClonable_Clone() As Object
    Set IClassClonable_Clone = Clone
End Function

Private Function ISupportSignal_Emitter() As SignalEmitter
    Set ISupportSignal_Emitter = mySignalEmitter
End Function

Private Property Get ISupportSlot_AutoManage() As Boolean
    ISupportSlot_AutoManage = False
End Property
Private Sub ISupportSlot_ManualManageCall(SlotName As String)
    Select Case SlotName
        Case SLOT_CLEAR
            Call Clear
        Case Else
            throw InvalidCallException
    End Select
End Sub
Private Function ISupportSlot_Methods() As String
    ISupportSlot_Methods = _
        "0" & SLOT_CLEAR
End Function

Private Property Get IAliasable_Alias() As String
    IAliasable_Alias = CLASSALIASSTRING
End Property
