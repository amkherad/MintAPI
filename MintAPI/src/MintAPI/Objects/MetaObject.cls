VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MetaObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"MetaObject"
Attribute VB_Ext_KEY = "Member1" ,"Stream"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member2" ,"Version"
Attribute VB_Ext_KEY = "Member3" ,"SignalEmitter"
'@PROJECT_LICENSE

Option Base 0
Option Explicit
'immutable
''<class>
''<summary>Provides meta-data for objects.</summary>
''<using>SignalEmitter,SerializedObject</using>
''<see>SerializedObject</see>
''</class>

'=============================================
'=============================================
'=============================================
'<section Meta Data>
'
Const CLASSID As String = "MetaObject"
Const CLASSALIAS As String = APPLICATIONDOMAIN & "." & CLASSID
Const CLASSVERSIONSTRING As String = "0.1"
Const CLASSALIASSTRING As String = CLASSALIAS & "/" & CLASSVERSIONSTRING

Const META_PROPERTY_SIGNALEMITTER As String = "signalemitter"
Const META_PROPERTY_SERIALIZER As String = "serializer"
Const META_PROPERTY_VERSION As String = "version"
'
'</section>
'---------------------------------------------
'---------------------------------------------
'---------------------------------------------

Implements IObject
Implements ICloneable

'=============================================
'=============================================
'=============================================
'<section API/Type Declarations>
'
'Private Declare Function API_IsBadCodePtr Lib "kernel32" Alias "IsBadCodePtr" (ByVal lpfn As Long) As Long
'
'</section>
'---------------------------------------------
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section Variables>
'
Dim This_Constructed As Boolean

Private p_Obj As New WeakReference
Private p_SignalEmitter As SignalEmitter
Private p_Alias As String
Private p_Version As Version
Private p_Synchronization As New LockHandle
Private p_ObjectFamilyChain As New List 'a list of objects that must be kept alive.
'
'</section>
'---------------------------------------------
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section Constructors>
'
Friend Sub Constructor(ByVal CInfo As ObjectConstructionInfo)
    
    This_Constructed = True
End Sub
Friend Sub Constructor0(ByVal tObject As Object, _
    ByVal SignalEmitter As Boolean, _
    ByVal AliasString As String)

    If Not TypeOf tObject Is IObject Then _
        throw Exps.InvalidArgumentException
    
    'Set p_Obj.Object = tObject
    p_Alias = AliasString
    If SignalEmitter Then
        Set p_SignalEmitter = New SignalEmitter
        Call p_SignalEmitter.Constructor0
    End If
    'Set tSerializer = Serializer
    
    This_Constructed = True
End Sub
Friend Sub ConstructorCopy(ByVal MetaObject As MetaObject)
    
    This_Constructed = True
End Sub
'
'</section>
'---------------------------------------------
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section Body Members>
'
Public Property Get SignalEmitter() As SignalEmitter
THROW_IfNull
    Set SignalEmitter = p_SignalEmitter
End Property
Public Property Get HaveSignals() As Boolean
THROW_IfNull
    HaveSignals = (Not (p_SignalEmitter Is Nothing))
End Property
'Public Property Get HaveSlots() As Boolean: THROW_IfNull
'
'End Property
'Public Function Slots() As Variant: THROW_IfNull
'
'End Function

''<summary>Retrieves the source object.</summary>
Public Property Get Object() As IObject
THROW_IfNull
    If p_Obj.IsAlive Then _
        Set Object = p_Obj.Object
End Property
''<summary>Represent object's name.</summary>
Public Property Get Name() As String
'#THROW_IfNull
    Name = TypeName(Object)
End Property

''<summary>Represents object's version.</summary>
Public Property Get Version() As Version
THROW_IfNull
    Version = p_Version
End Property

Public Property Get Stream() As IClassStream
THROW_IfNull
    Dim Obj As Object
    Set Obj = p_Obj.Object
    If Not TypeOf Obj Is IClassStream Then throw Exps.InvalidOperationException("Class Does Not Support Class Streaming.")
    Set Stream = Obj
End Property
Public Property Get SupportIO() As Boolean
THROW_IfNull
    Dim Obj As Object
    Set Obj = p_Obj.Object
    If TypeOf Obj Is IClassStream Then SupportIO = True
End Property

''<summary>Disposes all data and memory related to this instance of MetaObject.</summary>
Public Sub Dispose()
    Set p_Obj.Object = Nothing
    Set p_Version = Nothing
    If Not p_ObjectFamilyChain Is Nothing Then _
        Call p_ObjectFamilyChain.Clear
    Set p_ObjectFamilyChain = Nothing
    This_Constructed = False
End Sub

''<summary>Returns the object's interface information wrapper.</summary>
''<retval>An InterfaceInfo class to provide information about the object.</retval>
Public Property Get GetType() As CoClassInfo
THROW_IfNull
    Dim Obj As Object
    Set Obj = p_Obj.Object
    Set GetType = TypeInfo.ClassInfoFromObject(Obj)
End Property

Public Property Get Serializable() As Boolean
    
End Property


Public Function GetHashCode() As Long
THROW_IfNull
    Dim Obj As IObject
    Set Obj = p_Obj.Object
    GetHashCode = Obj.GetHashCode
End Function
''<summary>Returns a string which make object unique on the web.</summary>
''<retval>Object unique alias.</retval>
Public Property Get Alias() As String
THROW_IfNull
    Alias = p_Alias & " " & GetType().Guid.ToString
End Property
Public Function Equals(Value As Variant) As Boolean
THROW_IfNull
    Dim Obj As IObject
    Set Obj = p_Obj.Object
    Equals = Obj.Equals(Value)
End Function
Public Function MetaObject() As MetaObject
THROW_IfNull
    Set MetaObject = New MetaObject
    Call MetaObject.Constructor0(Me, False, CLASSVERSIONSTRING)
End Function
Private Function ToString() As String
THROW_IfNull
    Dim Obj As IObject
    Set Obj = p_Obj.Object
    ToString = Obj.ToString
End Function
''<summary>Provides class clone operation.</summary>
''<retval>A clone of the object instance.</retval>
''<sample>
''  <pre>Set Value = MetaObjectOf(Class).Clone</pre>
''</sample>
''<remarks>
''  Deprecated to use like <pre>MetaObjectOf(Class).Clone</pre> because
''  it takes too much performance.<br />
''  First to create MetaObject of object.<br />
''  Second to create a clone of object.<br />
''  Use Class.Clone instead.
''</remarks>
''<exceptions>ClassNotInitializedException,UnableToCloneObjectException</exception>
Public Function Clone() As Object
THROW_IfNull
    Dim Obj As Object
    Set Obj = p_Obj.Object
    If TypeOf Obj Is ICloneable Then
        Dim IClone As ICloneable
        Set IClone = Obj
        Set Clone = IClone.Clone
    Else
        throw Exps.InvalidCastException
    End If
End Function
''<summary>Check if object is cloneable.</summary>
''<retval>True if object is cloneable false if not.</retval>
''<exception>ClassNotInitializedException</exception>
Public Property Get Cloneable() As Boolean
THROW_IfNull
    Dim Obj As Object
    Set Obj = p_Obj.Object
    Cloneable = (TypeOf Obj Is ICloneable)
End Property


Friend Property Get Synchronization() As LockHandle
    Set Synchronization = p_Synchronization
End Property
Friend Property Get AliveChain() As Object
    Set AliveChain = p_ObjectFamilyChain
End Property

'
'</section>
'---------------------------------------------
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section IObject Implementations>
'
Private Sub IObject_Construct(ByVal CInfo As ObjectConstructionInfo): Call Constructor(CInfo): End Sub
Private Function IObject_Equals(Value As Variant) As Boolean: IObject_Equals = Me.Equals(Value): End Function
Private Function IObject_GetHashCode() As Long: IObject_GetHashCode = Me.GetHashCode: End Function
Private Function IObject_MetaObject() As MetaObject: Set IObject_MetaObject = Me.MetaObject: End Function
Private Function IObject_ToString() As String: IObject_ToString = CLASSALIASSTRING: End Function
Private Function IObject_Stub(ByVal ProxyInfo As ObjectProxyInfo) As Boolean
    
    IObject_Stub = False
End Function
'
'</section>
'---------------------------------------------
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section ICloneable Implementations>
'
Private Function ICloneable_Clone() As Object
    Dim MO As New MetaObject
    Call MO.ConstructorCopy(Me)
    Set ICloneable_Clone = MO
End Function
'
'</section>
'---------------------------------------------
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section Private Helpers>
'

'Private Sub Class_Terminate()
'
'End Sub

Private Sub THROW_IfNull()
    If (Not This_Constructed) Or (Not p_Obj.IsAlive) Then
        throw Exps.ClassNotInitializedException
    End If
End Sub
Private Sub Inner_ReadProperties(PropBag As PropertyBag, Optional IsClassEvent As Boolean = False)
    With PropBag
        Set p_Version = .ReadProperty(META_PROPERTY_VERSION, 1)
        'Set tSerializer = .ReadProperty(META_PROPERTY_SERIALIZER, Nothing)
        Set p_SignalEmitter = .ReadProperty(META_PROPERTY_SIGNALEMITTER, Nothing)
    End With
End Sub
Private Sub Inner_WriteProperties(PropBag As PropertyBag, Optional IsClassEvent As Boolean = False)
    With PropBag
        Call .WriteProperty(META_PROPERTY_VERSION, p_Version)
        'Call .WriteProperty(META_PROPERTY_SERIALIZER, tSerializer)
        Call .WriteProperty(META_PROPERTY_SIGNALEMITTER, p_SignalEmitter)
    End With
End Sub
'
'</section>
'---------------------------------------------
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section Class Events>
'
'Private Sub Class_Initialize()
'
'End Sub
'Private Sub Class_InitProperties()
'
'End Sub
'Private Sub Class_Terminate()
'
'End Sub
Private Sub Class_ReadProperties(PropBag As PropertyBag)
    Call Inner_ReadProperties(PropBag, True)
End Sub
Private Sub Class_WriteProperties(PropBag As PropertyBag)
    Call Inner_WriteProperties(PropBag, True)
End Sub
'
'</section>
'---------------------------------------------
'---------------------------------------------
'---------------------------------------------
