VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LockHandle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'@PROJECT_LICENSE

''<class name="LockHandle" tags="">
''  <summary>
''    Creates a simple mutual-exclusive lock object to be used in LockObj, TryLockObj and EndLockObj.
''  </summary>
''  <constructors>
''      <constructor name="" friend>
''          Some constructor description.
''      </constructor>
''  </constructors>
''  <using>Some class alias.</using>
''  <idea>Me</idea>
''  <sample>
''    Dim myLock As New LockHandle
''
''    LockObj myLock
''      'some statements.
''    EndLockObj myLock
''  </sample>
''</class>

Option Base 0
Option Explicit

Implements IObject
'Implements ICloneable
Implements IDisposable

'=============================================
'=============================================
'=============================================
'<section Meta Data>
'
Const CLASSID As String = "LockHandle" '#DO NOT REMOVE IT
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

Dim p_IsDisposed As Boolean
Dim p_Owner As Object
Dim p_SyncHandle As Long
Dim p_SyncHandleSpinCount As Long
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
'Friend Sub Constructor0(ByVal ID As Long)
'
'    This_Constructed = True
'End Sub
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

Friend Property Get SyncHandle() As Long
THROW_IfNull
    If p_SyncHandle = vbNullPtr Then
        Dim pSyncAccess As Long
        pSyncAccess = Interlocked.Increment(p_SyncHandleSpinCount)
        If (pSyncAccess <= 1) And (p_SyncHandle = vbNullPtr) Then
            p_SyncHandle = Memory.FastAllocate(API_CRITICALSECTION_SIZE)
            Call API_InitializeCriticalSection(p_SyncHandle)
        End If
        While p_SyncHandle = vbNullPtr: Wend
        Call Interlocked.Decrement(p_SyncHandleSpinCount)
    End If
    SyncHandle = p_SyncHandle
End Property

Public Property Get Owner() As Object
THROW_IfNull
    Set p_Owner = Owner
End Property

Public Sub BeginLock()
THROW_IfNull
    
End Sub
Public Function TryBeginLock() As Boolean
THROW_IfNull
    
End Function
Public Sub EndLock()
THROW_IfNull
    
End Sub

Public Property Get SyncCount() As Long
THROW_IfNull
    
End Property



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

Public Sub Dispose()
    p_IsDisposed = True
End Sub
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
    p_IsDisposed = False
End Sub
Private Sub Class_Terminate()
    If Not myMetaObject Is Nothing Then
        Signal.Emit Me, SIGNAL_TERMINATING
        Call myMetaObject.Dispose
        Set myMetaObject = Nothing
    End If
End Sub

Private Sub THROW_IfNull()
    If p_IsDisposed Then throw Exps.ObjectDisposedException
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
'Private Function ICloneable_Clone() As Object
'    Set ICloneable_Clone = Clone
'End Function
'
'</section>
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section IDisposable Implementations>
'
Private Sub IDisposable_Dispose(): Call Dispose: End Sub
'
'</section>
'---------------------------------------------
'---------------------------------------------
