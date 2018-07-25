VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MintAPIEventArgs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'@PROJECT_LICENSE

''<class name="" tags="" eventargs>
''  <summary>
''    Basic event argument.
''  </summary>
''  <sample></sample>
''</class>

Option Base 0
Option Explicit

Implements IObject
'Implements ICloneable
'Implements IDisposable

'=============================================
'=============================================
'=============================================
'<section Meta Data>
'
Const CLASSID As String = "MintAPIEventArgs" '#DO NOT REMOVE IT
Const CLASSALIAS As String = APPLICATIONDOMAIN & "." & CLASSID '#DO NOT REMOVE IT
Const CLASSVERSIONSTRING As String = "0.1" '#DO NOT REMOVE IT
Const CLASSALIASSTRING As String = CLASSALIAS & "/" & CLASSVERSIONSTRING '#DO NOT REMOVE IT

Const SIGNAL_TERMINATING                As String = "terminating"
Const SLOT_METAOBJECT                   As String = "metaobject"

Const PROPERTY_SENDER As String = "sender"
Const PROPERTY_VALUE As String = "value"
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

Private p_Sender As Object
Private p_Value As Variant
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
''  <param name="CInfo" type="ObjectConstructionInfo" byval in>
''      An object contains required info to construct the object.
''  </param>
''</params>
''<cargs order="0" count="1">
''  <carg type="long" name="ID"></carg>
''</cargs>
Friend Sub Constructor(ByVal CInfo As ObjectConstructionInfo)
    
    This_Constructed = True
End Sub
Friend Sub Constructor0(ByVal Sender As Object, ByVal Value As Variant)
    Set p_Sender = Sender
    Call Eval(p_Value, Value)
    
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

Public Property Get Sender() As Object: Set Sender = p_Sender: End Property

Public Property Get Value() As Variant: Set Value = p_Value: End Property

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

'Public Sub Dispose()
'
'End Sub
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

    'Call Dispose
End Sub

Private Sub Class_ReadProperties(PropBag As PropertyBag)
    With PropBag
'        Set p_Sender = .ReadProperty(PROPERTY_SENDER, Nothing)
        Call Eval(p_Value, .ReadProperty(PROPERTY_VALUE, False))
    End With
End Sub
Private Sub Class_WriteProperties(PropBag As PropertyBag)
    With PropBag
'        Call .WriteProperty(PROPERTY_SENDER, p_Sender, Nothing)
        Call .WriteProperty(PROPERTY_VALUE, p_Value, False)
    End With
End Sub
'
'</section>
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section EventArgs Implementation>
'
Private Function EventArgs_Clone() As EventArgs: Set EventArgs_Clone = Clone: End Function
Private Function EventArgs_Equals(Value As Variant) As Boolean: EventArgs_Equals = Equals(Value): End Function
Private Function EventArgs_GetHashCode() As Long: EventArgs_GetHashCode = GetHashCode: End Function
Private Function EventArgs_MetaObject() As MetaObject: Set EventArgs_MetaObject = MetaObject: End Function
Private Property Get EventArgs_Sender() As Object: Set EventArgs_Sender = Sender: End Property
Private Function EventArgs_ToString() As String: EventArgs_ToString = ToString: End Function
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
'Private Function ICloneable_Clone() As Object : Set ICloneable_Clone = Clone : End Function
'
'</section>
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section IDisposable Implementations>
'
'Private Sub IDisposable_Dispose(): Call Dispose: End Sub
'
'</section>
'---------------------------------------------
'---------------------------------------------
