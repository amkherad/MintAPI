VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MethodStatic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'@PROJECT_LICENSE

''<class name="MethodStatic" static tags="">
''  <summary>
''    Some class description.
''  </summary>
''  <using>Some class alias.</using>
''  <idea>Me</idea>
''  <sample></sample>
''</class>

Option Base 0
Option Explicit

'=============================================
'=============================================
'=============================================
'<section Meta Data>
'
Const CLASSID As String = "MethodStatic" '#DO NOT REMOVE IT
Const CLASSALIAS As String = APPLICATIONDOMAIN & "." & CLASSID '#DO NOT REMOVE IT
Const CLASSVERSIONSTRING As String = "0.1" '#DO NOT REMOVE IT
Const CLASSALIASSTRING As String = CLASSALIAS & "/" & CLASSVERSIONSTRING '#DO NOT REMOVE IT

'
'</section>
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section API/Type Declarations>
'

'
'</section>
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section Creators>
'
'ByVal Prototype As Prototype,
Public Function FromReference( _
        ByVal Name As String, _
        ByVal MethodReference As Long, _
        ByVal Prototype As MethodPrototype) As Method
'------------------------------------------
    If Prototype Is Nothing Then throw Exps.ArgumentNullException("Prototype")
    If MethodReference = vbNullPtr Then throw Exps.AccessDeniedException
    Set FromReference = New Method
    Call FromReference.Constructor0(Name, MethodReference, Prototype, False)
End Function
Friend Function FromReferenceUnsafe(ByVal Name As String, ByVal MethodReference As Long, _
        Optional ByVal CallingConvension As CallingConvensions = ccStdCall) As Method
    Set FromReferenceUnsafe = FromReference(Name, MethodReference, Prototype.UnspecifiedUnsafe)
End Function
''<default/>
Public Function Delegate(ByVal Obj As Object, ByVal MethodMember As String) As Method
Attribute Delegate.VB_UserMemId = 0
    Dim MemberInfo As MemberInfo
    Set MemberInfo = TypeInfo.ClassInfoFromObject(Obj).GetMember(MethodMember)
    
    If MemberInfo.MemberType <> mtMethod Then _
        throw Exps.InvalidCallException
    
    Set Delegate = MemberInfo.Method
End Function

Public Property Get Void() As Method
    Set Void = Method.FromReferenceUnsafe("Void", AddressOf Dummy_Void)
End Property
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

''<summary>Determines whether the method is create on the fly (runtime).</summary>
Public Function IsOnTheFly(ByVal Method As Method) As Boolean
    If Method Is Nothing Then throw Exps.ArgumentNullException("Method")
    IsOnTheFly = Method.IsOnTheFly
End Function
Public Function OnTheFlyLength(ByVal Method As Method) As Long
    If Method Is Nothing Then throw Exps.ArgumentNullException("Method")
    OnTheFlyLength = Method.OnTheFlyLength
End Function

Public Function DynamicInvoke(ByVal Obj As Object, ByVal Member As String, ByVal Args As ArgumentList, ByVal mType As MemberType) As Variant
    
End Function
Public Function ProfileInvoke(ByRef outRetVal As Variant, _
    ByVal Obj As Object, ByVal Member As String, ByVal Args As ArgumentList, _
    ByVal mType As MemberType, Optional ByVal ProfileMode As ProfilerModes = ProfilerModes.pmNone) As ProfilerInfo
    Set ProfileInvoke = New ProfilerInfo
    Call ProfileInvoke.StartProfile(ProfileMode)
    
    If IsMissing(outRetVal) Then
        Call DynamicInvoke(Obj, Member, Args, mType)
    Else
        Call MoveVariant(outRetVal, DynamicInvoke(Obj, Member, Args, mType))
    End If
    
    Call ProfileInvoke.EndProfile
End Function
Public Function ProfileInvokeAsync(ByRef outRetVal As Variant, _
    ByVal Obj As Object, ByVal Member As String, ByVal Args As ArgumentList, _
    ByVal mType As MemberType, Optional ByVal ProfileMode As ProfilerModes = ProfilerModes.pmNone) As ProfilerInfo
    Set ProfileInvokeAsync = New ProfilerInfo
    Call ProfileInvokeAsync.StartProfile(ProfileMode)
    
    If IsMissing(outRetVal) Then
        Call DynamicInvoke(Obj, Member, Args, mType)
    Else
        Call MoveVariant(outRetVal, DynamicInvoke(Obj, Member, Args, mType))
    End If
    
    Call ProfileInvokeAsync.EndProfile
End Function

Public Function BeginInvoke(ByVal Obj As Object, ByVal Member As String, ByVal Args As ArgumentList, ByVal mType As MemberType) As IAsyncResult
    
End Function
Public Sub EndInvoke(ByVal AR As IAsyncResult)
    
End Sub

Public Function ConcurrentInvoke(ByVal Obj As Object, ByVal Member As String, ByVal Args As ArgumentList, ByVal mType As MemberType) As Thread
    ConcurrentInvoke.Exec
End Function

Public Function IsExecutable(ByVal Reference As Long) As Boolean
    IsExecutable = (Reference <> 0)
    If IsExecutable Then _
        IsExecutable = (API_IsBadCodePtr(Reference) = NO_VALUE)
End Function

'*********************************************
'*********************************************
'
'</section>
'---------------------------------------------
'---------------------------------------------
