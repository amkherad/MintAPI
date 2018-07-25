VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Translation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'@PROJECT_LICENSE

''<class name="Translation" tags="">
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
Const CLASSID As String = "Translation" '#DO NOT REMOVE IT
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
Public Enum TranslationCacheState
    tcsOnMemory
    tcsOnDisk
    tcsAuto
End Enum
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

Private p_Cache As New Dictionary
Private p_CacheState As TranslationCacheState

Private p_Path As String
Private p_Name As String
Private p_IsoLetterSign As String
Private p_LCID As Long
Private p_Description As String
Private p_Copyright As String
Private p_Comment As String
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
Friend Sub Constructor0(ByVal Path As String)
    
    This_Constructed = True
End Sub
Friend Sub ConstructorCopy(ByVal That As Translation)
    
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

Public Property Get Name() As String
THROW_IfNull
    Name = p_Name
End Property
Public Property Get IsoLetterSign() As String
THROW_IfNull
    IsoLetterSign = p_IsoLetterSign
End Property
Public Property Get Culture() As Culture
THROW_IfNull
    Set Culture = MintAPI.Culture.FromLCID(p_LCID)
End Property
Public Property Get Description() As String
THROW_IfNull
    Description = p_Description
End Property
Public Property Get Copyright() As String
THROW_IfNull
    Copyright = p_Copyright
End Property
Public Property Get Comment() As String
THROW_IfNull
    Comment = p_Comment
End Property
Public Property Get Path() As String
THROW_IfNull
    Path = p_Path
End Property

Public Property Get CacheState() As TranslationCacheState
    CacheState = p_CacheState
End Property
Public Property Let CacheState(ByVal Value As TranslationCacheState)
    p_CacheState = Value
End Property


Public Function Translate(ByVal Key As String, Optional Default) As String
    If TryTranslate(Key, Translate) Then Exit Function
    If IsMissing(Default) Then
        throw Exps.ItemNotExistsException
    Else
        Translate = mString.Stringify(Default)
    End If
End Function
Public Function TryTranslate(ByVal Key As String, ByRef Value As String) As Boolean
    If TranslateFromMemory(Key, Value) Then TryTranslate = True: Exit Function
    If TranslateFromFile(Key, Value) Then TryTranslate = True: Exit Function
End Function

Public Function RecordExists(ByVal Key As String) As Boolean
    
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

Public Function Clone() As Translation
    Set Clone = New Translation
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

Private Function TranslateFromMemory(ByVal Key As String, ByRef Value As String) As Boolean
    Dim Val As Variant
    TranslateFromMemory = p_Cache.TryGetValue(Key, Val)
    Value = mString.Stringify(Val)
End Function
Private Function TranslateFromFile(ByVal Key As String, ByRef Value As String) As Boolean
    
End Function
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
