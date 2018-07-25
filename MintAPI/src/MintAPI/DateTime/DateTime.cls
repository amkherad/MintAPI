VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DateTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'@PROJECT_LICENSE

''<class name="DateTime" tags="">
''  <summary>
''    Represents date and time.
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
Implements ValueType

'=============================================
'=============================================
'=============================================
'<section Meta Data>
'
Const CLASSID As String = "DateTime" '#DO NOT REMOVE IT
Const CLASSALIAS As String = APPLICATIONDOMAIN & "." & CLASSID '#DO NOT REMOVE IT
Const CLASSVERSIONSTRING As String = "0.1" '#DO NOT REMOVE IT
Const CLASSALIASSTRING As String = CLASSALIAS & "/" & CLASSVERSIONSTRING '#DO NOT REMOVE IT

Const SIGNAL_TERMINATING                As String = "terminating"
Const SLOT_METAOBJECT                   As String = "metaobject"

Const PROPERTY_VALUE                    As String = "Value"
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

Dim p_Value As API_SYSTEMTIME
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
Friend Sub Constructor1(ByRef SystemTime As API_SYSTEMTIME)
    
    
    This_Constructed = True
End Sub
Friend Sub Constructor2(ByVal Value As Date)
    
    
    This_Constructed = True
End Sub
Friend Sub Constructor3(ByVal Year As Long, ByVal Month As Long, ByVal Day As Long, _
                        ByVal Hour As Long, ByVal Minute As Long, ByVal Second As Long, ByVal Millisecond As Long)
    
    This_Constructed = True
End Sub
Friend Sub ConstructorCopy(ByVal That As DateTime)

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

Public Function Add(ByVal TimeSpan As TimeSpan) As DateTime
    
End Function
Public Function AddYears(ByVal Years As Long) As DateTime
    
End Function
Public Function AddMonths(ByVal Months As Long) As DateTime
    
End Function
Public Function AddDays(ByVal Days As Long) As DateTime
    
End Function
Public Function AddHours(ByVal Hours As Long) As DateTime
    
End Function
Public Function AddMinutes(ByVal Minutes As Long) As DateTime
    
End Function
Public Function AddSeconds(ByVal Seconds As Long) As DateTime
    
End Function


Public Property Get Years() As Long
    Years = p_Value.wYear
End Property
Public Property Get Months() As Long
    Months = p_Value.wMonth
End Property
Public Property Get Days() As Long
    Days = p_Value.wDay
End Property
Public Property Get Hours() As Long
    Hours = p_Value.wHour
End Property
Public Property Get Minutes() As Long
    Minutes = p_Value.wMinute
End Property
Public Property Get Seconds() As Long
    Seconds = p_Value.wSecond
End Property
Public Property Get Milliseconds() As Long
    Milliseconds = p_Value.wMilliseconds
End Property


Public Property Get TotalYears() As Long
    
End Property
Public Property Get TotalMonths() As Long
    
End Property
Public Property Get TotalDays() As Long
    
End Property
Public Property Get TotalHours() As Long
    
End Property
Public Property Get TotalMinutes() As Long
    
End Property
Public Property Get TotalSeconds() As Long
    
End Property
Public Property Get TotalMilliseconds() As Long
    
End Property


Public Property Get SystemDateTime() As Date
    
End Property
Friend Property Let SystemDateTime(ByVal Value As Date)
    
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

Public Function Clone() As DateTime
    Set Clone = New DateTime
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

Private Sub Class_ReadProperties(PropBag As PropertyBag)
    SystemDateTime = PropBag.ReadProperty(PROPERTY_VALUE, Now)
End Sub
Private Sub Class_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty(PROPERTY_VALUE, SystemDateTime)
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
Private Function ICloneable_Clone() As Object: Set ICloneable_Clone = Clone: End Function
'
'</section>
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section ICloneable Implementations>
'
Private Property Get ValueType_Value() As Variant
    ValueType_Value = SystemDateTime
End Property
'
'</section>
'---------------------------------------------
'---------------------------------------------
