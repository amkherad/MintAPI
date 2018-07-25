VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SignalEmitter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"MetaObject"
'@PROJECT_LICENSE

''<class name="SignalEmitter" tags="">
''  <summary>
''    Some class description.
''  </summary>
''  <constructors>
''      <constructor name="" friend>
''          Some constructor description.
''      </constructor>
''  </constructors>
''  <using>Some class alias.</using>
''  <idea>Me</idea>
''  <sample></sample>
''</class>

Option Base 0
Option Explicit

Implements IObject
Implements ICloneable

'=============================================
'=============================================
'=============================================
'<section Meta Data>
'
Const CLASSID As String = "SignalEmitter" '#DO NOT REMOVE IT
Const CLASSALIAS As String = APPLICATIONDOMAIN & "." & CLASSID '#DO NOT REMOVE IT
Const CLASSVERSIONSTRING As String = "0.1" '#DO NOT REMOVE IT
Const CLASSALIASSTRING As String = CLASSALIAS & "/" & CLASSVERSIONSTRING '#DO NOT REMOVE IT

Const SIGNAL_TERMINATING                As String = "terminating"
Const SLOT_METAOBJECT                   As String = "metaobject"
'
'</section>
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section API/Type Declarations>
'
Private Declare Function API_IsBadCodePtr Lib "kernel32" Alias "IsBadCodePtr" (ByVal lpfn As Long) As Long


Private Enum SlotType
    stMethod
    stCallBack
    stObject
    stAddress
End Enum
Private Enum CallingMethodType
    cmtLower = 0
    cmtManual = 0
    cmtAuto = 1
    cmtNotExists = 2
    cmtUpper = 2
End Enum

Private Type SlotHolder
    tType As SlotType
    tMethod As Method
    tCallBack As ICallBack
    tObject As Object
    tAddress As Long
    SlotName As String
End Type
Private Type SignalBounds
    SignalName As String
    Slots() As SlotHolder
    slotsCount As Long
    State As Boolean
End Type
'
'</section>
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section Variables/Events>
'
Dim This_Constructed As Boolean
Dim myMetaObject As MetaObject

Dim Sig As IObject
Dim OriginalSE As SignalEmitter

Dim Sigs As New Collection 'List
Dim SigState As Boolean
'
'</section>
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section Constructors>
'
''<summary>
''  Basic constructor to use by MintAPI class provider,
''  When creating new instance of this class.
''</summary>
''<params>
''  <param name="Arguments">
''      A list of construction arguments.
''  </param>
''</params>
''<cargs order="0" count="1">
''  <carg type="long" name="ID"></carg>
''</cargs>
Friend Sub Constructor(ByVal CInfo As ObjectConstructionInfo)
    
    This_Constructed = True
End Sub
Friend Sub Constructor0()
    
    This_Constructed = True
End Sub
Friend Sub ConstructorCopy(ByVal SignalEmitter As SignalEmitter)
    
    This_Constructed = True
End Sub
'
'</section>
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section Class Public Body Members>
'
'*********************************************
'*********************************************

Public Property Get Object() As IObject
    Set Object = Sig
End Property

Public Function Emit(Signal, Optional ByVal Args As ArgumentList = Nothing) As Variant
    
End Function

Public Sub EmitAll()
    Dim Obj As Object, Method As Method
    For Each Obj In Sigs
        Set Method = Obj
        Call InvokeFor(Method, Nothing)
    Next
End Sub


Public Function ExternEmitter() As SignalEmitter
    Set ExternEmitter = Clone
    Call ExternEmitter.SetOriginalSE(Me)
End Function

'*********************************************
'*********************************************
'
'</section>
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section Basic Members>
'
Public Function Equals(Value As Variant) As Boolean
    Equals = Objects.Equals(Me, Value)
End Function
Public Function ToString() As String
    ToString = CLASSALIAS
End Function
Public Function GetHashCode() As Long
    GetHashCode = Objects.GetHashCode(Me)
End Function
Public Function MetaObject() As MetaObject
    If myMetaObject Is Nothing Then
        Call Objects.MetaObject.Setup(myMetaObject, Me, True, CLASSALIASSTRING)
    End If
    Set MetaObject = myMetaObject
End Function

Public Function Clone() As SignalEmitter
    Set Clone = New SignalEmitter
    Call Clone.ConstructorCopy(Me)
End Function
'
'</section>
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section Private helpers>
'
Private Sub Class_Initialize()
    SigState = True
End Sub
Private Sub Class_Terminate()
    Set Sig = Nothing
    Set OriginalSE = Nothing
    Set Sigs = Nothing
    If Not myMetaObject Is Nothing Then
        Call myMetaObject.Dispose
        Set myMetaObject = Nothing
    End If
End Sub

Friend Sub SetOriginalSE(ByVal OriginalSignalEmitter As SignalEmitter)
    Set OriginalSE = OriginalSignalEmitter
End Sub

Private Function InvokeFor(ByVal Method As Method, ByVal Args As ArgumentList)
    
End Function

Friend Property Get SignalsState() As Boolean
    SignalsState = SigState
End Property
Friend Property Let SignalsState(ByVal Value As Boolean)
    SigState = Value
End Property

Friend Property Get SignalState(ByVal SignalName As String) As Boolean
    
End Property
Friend Property Let SignalState(ByVal SignalName As String, ByVal Value As Boolean)
    
End Property

'Private Function IsTrueSlotObjectBuffer(SlotObjectBuffer As ObjectBuffer) As Boolean
'    Dim ArgsCount As Long
'    ArgsCount = SlotObjectBuffer.Count
'    If ArgsCount = 1 Then
'        Dim ArgType As VbVarType
'        ArgType = SlotObjectBuffer.ArgType(0)
'        Select Case ArgType
'            Case vbLong
'                If CBool(API_IsBadCodePtr(SlotObjectBuffer.First)) Then Exit Function
'            Case VBObject
'                If Not _
'                   ((TypeOf SlotObjectBuffer.First Is ICallBack) Or _
'                    (TypeOf SlotObjectBuffer.First Is Method)) Then Exit Function
'        End Select
'    ElseIf ArgsCount = 2 Then
'        If SlotObjectBuffer.ArgType(0) <> vbString Then Exit Function
'        If SlotObjectBuffer.ArgType(1) <> VBObject Then Exit Function
'        If SlotObjectBuffer.Last Is Nothing Then Exit Function
'        If Not (TypeOf SlotObjectBuffer.Last Is ISupportSlot) Then Exit Function
'    Else
'        Exit Function
'    End If
'    IsTrueSlotObjectBuffer = True
'End Function
'Private Function GetSlotHolderOf(SlotObjectBuffer As ObjectBuffer, Optional ThrowNotObject As Boolean = False) As SlotHolder
'    Dim ArgsCount As Long
'    ArgsCount = SlotObjectBuffer.Count
'    If ArgsCount = 1 Then
'        Dim ArgType As VbVarType
'        ArgType = SlotObjectBuffer.ArgType(0)
'        Select Case ArgType
'            Case vbLong
'                GetSlotHolderOf.tType = stAddress
'                GetSlotHolderOf.tAddress = CLng(SlotObjectBuffer.First)
'            Case VBObject
'                If (TypeOf SlotObjectBuffer.First Is ICallBack) Then
'                    GetSlotHolderOf.tType = stCallBack
'                    Set GetSlotHolderOf.tCallBack = SlotObjectBuffer.First
'                ElseIf (TypeOf SlotObjectBuffer.First Is Method) Then
'                    GetSlotHolderOf.tType = stMethod
'                    Set GetSlotHolderOf.tMethod = SlotObjectBuffer.First
'                Else
'                    throw Exps.InvalidArgumentTypeException("at SignalEmitter.GetSlotHolderOf().")
'                End If
'        End Select
'    ElseIf ArgsCount = 2 Then
'        GetSlotHolderOf.tType = stObject
'        GetSlotHolderOf.SlotName = SlotObjectBuffer.First
'        Set GetSlotHolderOf.tObject = SlotObjectBuffer.Last
'        If ThrowNotObject Then
'            If GetSlotHolderOf.tObject Is Nothing Then _
'                throw Exps.ArgumentNullException
'        End If
'    Else
'        throw Exps.InvalidArgumentTypeException("at SignalEmitter.GetSlotHolderOf().")
'    End If
'End Function
'
'
'Private Function CompareSlotHolders(SH1 As SlotHolder, SH2 As SlotHolder) As Long
'    CompareSlotHolders = -1
'    If SH1.tType <> SH2.tType Then Exit Function
'
'    Select Case SH1.tType
'        Case stObject
'            If LCase(SH1.SlotName) = LCase(SH2.SlotName) Then _
'                If SH1.tObject Is SH2.tObject Then CompareSlotHolders = 1
'        Case stAddress
'            If SH1.tAddress = SH2.tAddress Then CompareSlotHolders = 1
'        Case stMethod
'            If SH1.tMethod.Reference = SH2.tMethod.Reference Then CompareSlotHolders = 1
'        Case stCallBack
'            If SH1.tCallBack Is SH2.tCallBack Then CompareSlotHolders = 1
'        Case Else
'            throw Exps.InvalidStatusException
'    End Select
'End Function
'
'Private Function IndexOf(ByVal SignalName As String) As Long
'    Dim i As Long
'    SignalName = LCase$(SignalName)
'    For i = 0 To sCount - 1
'        If S(i).SignalName = SignalName Then
'            IndexOf = i
'            Exit Function
'        End If
'    Next
'    IndexOf = -1
'End Function
'
'Friend Sub Append(SignalName As String, SlotObjectBuffer As ObjectBuffer)
'    If Not OriginalSE Is Nothing Then
'        Call OriginalSE.Append(SignalName, SlotObjectBuffer)
'        Exit Sub
'    End If
'    If SlotObjectBuffer Is Nothing Then throw Exps.ArgumentNullException
'    If Not IsTrueSlotObjectBuffer(SlotObjectBuffer) Then _
'        throw Exps.InvalidArgumentTypeException("SlotObjectBuffer does not meet the standard requirements of slot object buffer.")
'
'    Dim Index As Long, sIndex As Long
'    Dim SH As SlotHolder
'    Index = IndexOf(SignalName)
'
'    If Index = -1 Then
'        SH = GetSlotHolderOf(SlotObjectBuffer)
'        ReDim Preserve S(sCount)
'        S(sCount).SignalName = SignalName
'        ReDim S(sCount).Slots(0)
'        S(sCount).slotsCount = 1
'        S(sCount).Slots(0) = SH
'        S(sCount).State = True
'        sCount = sCount + 1
'    Else
'        SH = GetSlotHolderOf(SlotObjectBuffer)
'        sIndex = S(Index).slotsCount
'        ReDim Preserve S(Index).Slots(sIndex)
'        S(Index).slotsCount = S(Index).slotsCount + 1
'        S(Index).Slots(sIndex) = SH
'    End If
'End Sub
'Friend Sub Remove(SignalName As String, SlotObjectBuffer As ObjectBuffer)
'    If Not OriginalSE Is Nothing Then
'        Call OriginalSE.Remove(SignalName, SlotObjectBuffer)
'        Exit Sub
'    End If
'    If SlotObjectBuffer Is Nothing Then throw Exps.ArgumentNullException
'    If Not IsTrueSlotObjectBuffer(SlotObjectBuffer) Then _
'        throw Exps.InvalidArgumentTypeException("SlotObjectBuffer does not meet the standard requirements of slot object buffer.")
'
'    Dim Index As Long
'    Index = IndexOf(SignalName)
'
'    If Index = -1 Then throw Exps.OutOfRangeException("Target signal not found.")
'
'    Dim i As Long, Cmp As Long, SH As SlotHolder, j As Long
'    SH = GetSlotHolderOf(SlotObjectBuffer)
'    For i = 0 To S(Index).slotsCount - 1
'        Cmp = CompareSlotHolders(S(Index).Slots(i), SH)
'        If Cmp = 1 Then
'            For j = i To S(Index).slotsCount - 2
'                S(Index).Slots(j) = S(Index).Slots(j + 1)
'            Next
'            S(Index).slotsCount = S(Index).slotsCount - 1
'            If S(Index).slotsCount > 0 Then
'                ReDim Preserve S(Index).Slots(S(Index).slotsCount - 1)
'            Else
'                Erase S(Index).Slots
'            End If
'            Exit For
'        End If
'    Next
'End Sub
'Friend Sub ClearSignal(SignalName As String)
'    If Not OriginalSE Is Nothing Then
'        Call OriginalSE.ClearSignal(SignalName)
'        Exit Sub
'    End If
'    Dim Index As Long
'    Index = IndexOf(SignalName)
'
'    If Index = -1 Then throw Exps.OutOfRangeException("Target signal not found.")
'
'    Dim i As Long
'    For i = Index To sCount - 2
'        S(i) = S(i + 1)
'    Next
'    sCount = sCount - 1
'    If sCount > 0 Then
'        ReDim Preserve S(sCount - 1)
'    Else
'        Erase S
'    End If
'End Sub
'Friend Sub ClearSlotOf(ClassName As String)
'    If Not OriginalSE Is Nothing Then
'        Call OriginalSE.ClearSlotOf(ClassName)
'        Exit Sub
'    End If
'
'    Dim xi As Long, i As Long, Cmp As Long, j As Long
'    For xi = 0 To sCount - 1
'        For i = 0 To S(xi).slotsCount - 1
'            If S(xi).Slots(i).tType = stObject Then
'                If TypeName(S(xi).Slots(i).tObject) = ClassName Then
'                    For j = i To S(xi).slotsCount - 2
'                        S(xi).Slots(j) = S(xi).Slots(j + 1)
'                    Next
'                    S(xi).slotsCount = S(xi).slotsCount - 1
'                    If S(xi).slotsCount > 0 Then
'                        ReDim Preserve S(xi).Slots(S(xi).slotsCount - 1)
'                    Else
'                        Erase S(xi).Slots
'                    End If
'                    Exit Sub
'                End If
'            End If
'        Next
'    Next
'End Sub
'
'Private Function GetCallingMethodType(strSource As String, strMethod As String) As CallingMethodType
'    Dim i As Long
'    Dim chChar As String * 1
'    Dim chByte As Byte
'    Dim splStr() As String
'    splStr = Split(strSource, " ")
'
'    For i = 0 To ArraySize(splStr) - 1
'        If splStr(i) = strMethod Then
'            GetCallingMethodType = cmtManual
'            Exit Function
'        ElseIf Right(splStr(i), Len(splStr(i)) - 1) = strMethod Then
'            chChar = Mid(splStr(i), 1, 1)
'            chByte = Asc(chChar)
'            If chByte >= 48 And chByte <= 57 Then
'                GetCallingMethodType = chByte - 48
'                If GetCallingMethodType < cmtLower Or _
'                   GetCallingMethodType > cmtUpper Then
'                   GetCallingMethodType = cmtManual
'                End If
'                Exit Function
'            Else
'                GetCallingMethodType = cmtManual
'                Exit Function
'            End If
'        End If
'    Next
'
'    GetCallingMethodType = cmtNotExists
'End Function
'Private Sub EmitISupportSlot(SignalName As String, ISSO As Object, SlotName As String)
'    If ((ISSO Is Nothing) Or (Not TypeOf ISSO Is ISupportSlot)) Then _
'        throw Exps.ArgumentNullException("Requires a ISupportSlot class.")
'    Dim ISS As ISupportSlot
'    Set ISS = ISSO
'    If ISS.AutoManage Then
'        Dim CMT As CallingMethodType
'        CMT = GetCallingMethodType(ISS.Methods, SlotName)
'        Select Case CMT
'            Case cmtAuto
'                On Error GoTo callbyname_Error
'                CallByName ISSO, SlotName, VbMethod
'                GoTo dropError
'callbyname_Error:
'                throw Exps.InvalidCallException("Invalid Slot call.")
'dropError:
'            Case cmtManual
'                Call ISS.ManualManageCall(LCase(SlotName))
'            Case cmtNotExists
'                throw Exps.Exception("Target Slot does not exists in target class.")
'        End Select
'    Else
'        Call ISS.ManualManageCall(LCase(SlotName))
'    End If
'End Sub
'Private Sub EmitMethod(SignalName As String, Method As Method, SlotName As String)
'
'End Sub
'Private Sub EmitCallBack(ByVal SignalName As String, Callback As ICallBack, SlotName As String)
'    'Call callback.Invoke(TypeName(sig) & " has emitted " & SignalName & ".")
'End Sub
'Private Sub EmitAddress(SignalName As String, Address As Long, SlotName As String)
'    'Call Method(SlotName, Address).Invoke
'End Sub
'
'Private Sub EmitSlotHolder(SignalName As String, SH As SlotHolder)
'    Select Case SH.tType
'        Case stMethod
'            Call EmitMethod(SignalName, SH.tMethod, SH.SlotName)
'        Case stCallBack
'            Call EmitCallBack(SignalName, SH.tCallBack, SH.SlotName)
'        Case stObject
'            Call EmitISupportSlot(SignalName, SH.tObject, SH.SlotName)
'        Case stAddress
'            Call EmitAddress(SignalName, SH.tAddress, SH.SlotName)
'        Case Else
'            throw Exps.InvalidStatusException
'    End Select
'End Sub
'
''Public Sub EmitExpress(SignalName As String)
''
''End Sub
'Friend Function EmitC(Signal, Args() As Variant) As Variant
'    If Not SigState Then Exit Function
'    If Not OriginalSE Is Nothing Then
'        Call OriginalSE.Emit(Signal)
'        Exit Function
'    End If
'    If sCount <= 0 Then Exit Function
'    Dim SignalName As String
'    Dim Index As Long
'
'    If IsObject(Signal) Then
'        If TypeOf Signal Is ObjectBuffer Then
'            Dim OB As ObjectBuffer
'            Set OB = Signal
'            Call OB.ThrowNot(OBJECTBUFFER_SIGNAL, 1)
'            SignalName = CStr(OB.First)
'        Else
'            throw Exps.InvalidVarTypeException("at SignalEmitter.Emit() , Invalid object.")
'        End If
'    Else
'        SignalName = CStr(Signal)
'    End If
'
'    Index = IndexOf(SignalName)
'    If Index = -1 Then Exit Function
'
'    If Not S(Index).State Then Exit Function
'
'    Dim i As Long
'    For i = 0 To S(Index).slotsCount - 1
'        Call EmitSlotHolder(SignalName, S(Index).Slots(i))
'    Next
'End Function
'
'</section>
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section IObject Implementation>
'
Private Sub IObject_Construct(ByVal CInfo As ObjectConstructionInfo): Call Constructor(CInfo): End Sub
Private Function IObject_Equals(Value As Variant) As Boolean: IObject_Equals = Equals(Value): End Function
Private Function IObject_GetHashCode() As Long: IObject_GetHashCode = GetHashCode: End Function
Private Function IObject_MetaObject() As MetaObject: Set IObject_MetaObject = MetaObject: End Function
Private Function IObject_ToString() As String: IObject_ToString = ToString: End Function
Private Function IObject_Stub(ByVal ProxyInfo As ObjectProxyInfo) As Boolean
    
    IObject_Stub = False
End Function
'
'</section>
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section ICloneable Implementations>
'
Private Function ICloneable_Clone() As Object: Set ICloneable_Clone = Clone: End Function
'
'</section>
'---------------------------------------------
'---------------------------------------------
