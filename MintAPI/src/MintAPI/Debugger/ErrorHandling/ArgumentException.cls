VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ArgumentException"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@PROJECT_LICENSE

''<class name="ArgumentException" tags="">
''  <summary>
''    Represents a basic argument exception.
''  </summary>
''  <using>Some class alias.</using>
''  <idea>Me</idea>
''  <sample></sample>
''</class>

Option Base 0
Option Explicit

Implements IObject
Implements ICloneable
Implements Exception

'=============================================
'=============================================
'=============================================
'<section Meta Data>
'
Const CLASSID As String = "ArgumentException" '#DO NOT REMOVE IT
Const CLASSALIAS As String = APPLICATIONDOMAIN & "." & CLASSID '#DO NOT REMOVE IT
Const CLASSVERSIONSTRING As String = "0.1" '#DO NOT REMOVE IT
Const CLASSALIASSTRING As String = CLASSALIAS & "/" & CLASSVERSIONSTRING '#DO NOT REMOVE IT

Const SIGNAL_TERMINATING As String = "terminating"
Const SLOT_METAOBJECT As String = "metaobject"

Const PROPERTY_MESSAGE As String = "message"
Const PROPERTY_DETAILS As String = "details"
Const PROPERTY_INNEREXCEPTION As String = "innerexception"
Const PROPERTY_HELPLINK As String = "helplink"
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

Dim p_Message As String
Dim p_Details As String
Dim p_InnerException As Exception
Dim p_HelpLink As String
Dim p_tObject As Object
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
Friend Sub Constructor0( _
            ByVal Message As String, _
            ByVal InnerException As Exception)
    
    p_Message = Message
    Set p_InnerException = InnerException
    
    This_Constructed = True
End Sub
Friend Sub ConstructorCopy(ByVal That As Exception)
    p_Message = That.Message
    p_Details = That.Details
    p_HelpLink = That.HelpLink
    Set p_InnerException = That.InnerException
    Set p_tObject = That.Object
    
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

Public Property Get Message() As String
    Message = p_Message
End Property
Public Property Get Details() As String
    Details = p_Details
End Property
Public Property Let Details(ByVal Value As String)
    p_Details = Value
End Property
Public Property Get InnerException() As Exception
    Set InnerException = p_InnerException
End Property
Public Property Get HelpLink() As String
    HelpLink = p_HelpLink
End Property
Public Property Let HelpLink(ByVal Value As String)
    p_HelpLink = Value
End Property

Public Property Get Object() As Object
    Set Object = p_tObject
End Property
Public Property Let Object(ByVal Obj As Object)
    Set p_tObject = Obj
End Property
Public Property Set Object(ByVal Obj As Object)
    Set p_tObject = Obj
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

Public Function Clone() As Object
    Dim Exp As New Exception
    Call Exp.ConstructorCopy(Me)
    Set Clone = Exp
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
    p_Message = ERRORS_ARGUMENTEXCEPTION
End Sub
Private Sub Class_Terminate()
    If Not myMetaObject Is Nothing Then
        Signal.Emit Me, SIGNAL_TERMINATING
        Call myMetaObject.Dispose
        Set myMetaObject = Nothing
    End If
End Sub

Private Sub Class_ReadProperties(PropBag As PropertyBag)
    With PropBag
        p_Message = .ReadProperty(PROPERTY_MESSAGE, "")
        p_Details = .ReadProperty(PROPERTY_DETAILS, "")
        p_InnerException = .ReadProperty(PROPERTY_INNEREXCEPTION, "")
        p_HelpLink = .ReadProperty(PROPERTY_HELPLINK, "")
    End With
End Sub
Private Sub Class_WriteProperties(PropBag As PropertyBag)
    With PropBag
        Call .WriteProperty(PROPERTY_MESSAGE, p_Message)
        Call .WriteProperty(PROPERTY_DETAILS, p_Details)
        Call .WriteProperty(PROPERTY_INNEREXCEPTION, p_InnerException)
        Call .WriteProperty(PROPERTY_HELPLINK, p_HelpLink)
    End With
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
Private Function ICloneable_Clone() As Object
    Set ICloneable_Clone = Clone
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
Private Function Exception_Clone() As Object: Set Exception_Clone = Clone: End Function
Private Property Let Exception_Details(ByVal RHS As String): Details = RHS: End Property
Private Property Get Exception_Details() As String: Exception_Details = Details: End Property
Private Function Exception_Equals(Value As Variant) As Boolean: Exception_Equals = Equals(Value): End Function
Private Function Exception_GetHashCode() As Long: Exception_GetHashCode = GetHashCode: End Function
Private Property Let Exception_HelpLink(ByVal RHS As String): HelpLink = RHS: End Property
Private Property Get Exception_HelpLink() As String: Exception_HelpLink = HelpLink: End Property
Private Property Get Exception_InnerException() As Exception: Set Exception_InnerException = InnerException: End Property
Private Property Get Exception_Message() As String: Exception_Message = Message: End Property
Private Function Exception_MetaObject() As MetaObject: Set Exception_MetaObject = MetaObject: End Function
Private Property Get Exception_Object() As Object: Set Exception_Object = Object: End Property
Private Property Let Exception_Object(ByVal Value As Object): Set Object = Value: End Property
Private Property Set Exception_Object(ByVal Value As Object): Set Object = Value: End Property
Private Function Exception_ToString() As String: Exception_ToString = ToString: End Function
'
'</section>
'---------------------------------------------
'---------------------------------------------
