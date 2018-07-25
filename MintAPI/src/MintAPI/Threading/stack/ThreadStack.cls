VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ThreadStack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"MetaObject"
'@PROJECT_LICENSE

''<class name="ThreadStack" tags="proxy_immutable">
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
'Implements ICloneable

'=============================================
'=============================================
'=============================================
'<section Meta Data>
'
Const CLASSID As String = "ThreadStack" '#DO NOT REMOVE IT
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

Dim p_StackPtr As Long
Dim p_Bounds As StackBoundries
Dim p_Thread As Thread
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
Friend Sub Constructor0(ByVal StackPtr As Long, ByVal Thread As Thread, ByRef Bounds As StackBoundries)
    p_StackPtr = StackPtr
    Set p_Thread = Thread
    p_Bounds = Bounds
    
    This_Constructed = True
End Sub
'Friend Sub ConstructorCopy(ByVal That As {My Object Type})
'
'    This_Constructed = True
'End Sub
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

Public Property Get LowerBound() As Long
THROW_IfNull
    LowerBound = p_Thread.StackBounderies.StackBottom
End Property
Public Property Get UpperBound() As Long
THROW_IfNull
    UpperBound = p_Thread.StackBounderies.StackTop
End Property

Public Property Get Size() As Long
THROW_IfNull
    Dim Bounds As StackBoundries
    Bounds = p_Thread.StackBounderies
    Size = (Bounds.StackTop - Bounds.StackBottom)
End Property

Public Property Get StackPointer() As Long
THROW_IfNull
    
End Property
Private Property Let StackPointer(ByVal Value As Long)
THROW_IfNull
    
End Property

Public Property Get Thread() As Thread
THROW_IfNull
    Set Thread = p_Thread
End Property

Public Property Get IsCurrentStack() As Boolean
THROW_IfNull
    IsCurrentStack = p_Thread.IsCurrentThread
End Property

''<summary>
''  Push'es the data to thread's stack.
''</summary>
''<remarks>
''  The caller is responsible to Pop the data from stack,
''  otherwise the calling method can't return and this cause process to crash.
''</remarks>
Public Sub Push(ByVal Buffer As Buffer)
THROW_IfNull
    Dim TEBPtr As Long, EBP As Long, ESP As Long
    
    If IsCurrentStack Then
        throw Exps.InvalidOperationException
    End If
    'TEBPtr = Thread.GetFS18
    
End Sub
''<summary>
''  Pop's the data from thread's stack, the data will be removed from stack.
''<summary>
''<remarks>
''  The caller is responsible to calling method's return pointer not to be removed,
''  otherwise the calling method can't return and this cause process to crash.
''<remarks>
Public Function Pop(ByVal Length As Long) As Buffer
THROW_IfNull
    Dim TEBPtr As Long, EBP As Long, ESP As Long
    
    If IsCurrentStack Then
        throw Exps.InvalidOperationException
    End If
    'TEBPtr = Thread.GetFS18
    
End Function
''<summary>
''  Peek's the data from thread's stack, the source data in stack will be untouched.
''<summary>
''<remarks>
''  This function is optimized for Thread to be Nothing, don't pass Thread.Current.
''<remarks>
''<exceptions>
''  <exception name="Exception">This will be thrown when ReadType length and Length have conflicts.</exception>
''</exceptions>
Public Function Peek(ByVal Length As Long) As Buffer
THROW_IfNull
    Dim TEBPtr As Long, EBP As Long, ESP As Long
    
    If IsCurrentStack Then
        throw Exps.InvalidOperationException
    End If
    'TEBPtr = Thread.GetFS18
    
End Function
''<summary>
''  Read's the data from thread's stack.
''<summary>
''<exceptions>
''  <exception name="Exception">This will be thrown when ReadType length and Length have conflicts.</exception>
''</exceptions>
Public Sub ReadData(ByRef Bytes() As Byte, ByVal Offset As Long, ByVal Length As Long)
THROW_IfNull
    Dim TEBPtr As Long, EBP As Long, ESP As Long
    
    If IsCurrentStack Then
        throw Exps.InvalidOperationException
    End If
    'TEBPtr = Thread.GetFS18
    
End Sub
''<summary>
''  Write's the data into thread's stack.
''<summary>
''<exceptions>
''  <exception name="Exception">This will be thrown when ReadType length and Length have conflicts.</exception>
''</exceptions>
Public Sub WriteData(ByRef Bytes() As Byte, ByVal Offset As Long, ByVal Length As Long)
THROW_IfNull
    Dim TEBPtr As Long, EBP As Long, ESP As Long
    
    If IsCurrentStack Then
        throw Exps.InvalidOperationException
    End If
    'TEBPtr = Thread.GetFS18
    
End Sub

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

'Public Function Clone() As Object
'
'End Function
'
'</section>
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section Private helpers>
'
'Private Sub Class_Initialize()
'    ...
'End Sub
Private Sub Class_Terminate()
    If Not myMetaObject Is Nothing Then
        Signal.Emit Me, SIGNAL_TERMINATING
        Call myMetaObject.Dispose
        Set myMetaObject = Nothing
    End If
End Sub

Private Sub THROW_IfNull()
    
End Sub
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
'Private Function ICloneable_Clone() As Object: Set ICloneable_Clone = Clone: End Function
'
'</section>
'---------------------------------------------
'---------------------------------------------
