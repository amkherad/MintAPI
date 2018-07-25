VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MenuItemEventArgs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'@PROJECT_LICENSE

''<class name="MenuItemEventArgs" tags="">
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
Implements EventArgs

'=============================================
'=============================================
'=============================================
'<section Meta Data>
'
Const CLASSID As String = "MenuItemEventArgs" '#DO NOT REMOVE IT
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

Dim p_Sender As Object
Dim p_MenuItem As MenuItem
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
Friend Sub ConstructorArg(ByVal Arguments As ArgumentList)
    
    This_Constructed = True
End Sub
Friend Sub Constructor0(ByVal Sender As Object, ByVal MenuItem As MenuItem)
    Set p_Sender = Sender
    Set p_MenuItem = MenuItem
    
    This_Constructed = True
End Sub
Friend Sub ConstructorCopy(ByVal That As MenuItemEventArgs)
    Set p_Sender = That.Sender
    Set p_MenuItem = That.MenuItem
    
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

Public Property Get Sender() As Object
    Set Sender = p_Sender
End Property

Public Property Get MenuItem() As MenuItem
    Set MenuItem = p_MenuItem
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

Public Function Clone() As MenuItemEventArgs
    Set Clone = New MenuItemEventArgs
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
'
'</section>
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section IObject Implementation>
'
Private Sub IObject_Initialize(ByVal Arguments As ArgumentList)
    Call ConstructorArg(Arguments)
End Sub
Private Function IObject_Equals(Value As Variant) As Boolean
    IObject_Equals = Equals(Value)
End Function
Private Function IObject_GetHashCode() As Long
    IObject_GetHashCode = GetHashCode
End Function
Private Function IObject_MetaObject() As MetaObject
    Set IObject_MetaObject = MetaObject
End Function
Private Function IObject_ToString() As String
    IObject_ToString = ToString
End Function
Private Function IObject_QueryInformation(ByVal OQR As ObjectQR) As Boolean
    
    IObject_QueryInformation = False
End Function
Private Function IObject_Notify(ByVal Name As String, ByVal Arguments As ArgumentList) As Variant
    Select Case LCase$(Name)
        'Case ""
            
        Case Else
            throw Exps.UnknownNotificationException, Me
    End Select
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
'<section EventArgs Implementations>
'
Private Function EventArgs_Clone() As MintAPI.EventArgs: Set EventArgs_Clone = Clone: End Function
Private Function EventArgs_Equals(Value As Variant) As Boolean: EventArgs_Equals = Equals(Value): End Function
Private Function EventArgs_GetHashCode() As Long: EventArgs_GetHashCode = GetHashCode: End Function
Private Function EventArgs_MetaObject() As MintAPI.MetaObject: Set EventArgs_MetaObject = MetaObject: End Function
Private Property Get EventArgs_Sender() As Object: Set EventArgs_Sender = Sender: End Property
Private Function EventArgs_ToString() As String: EventArgs_ToString = ToString: End Function
'
'</section>
'---------------------------------------------
'---------------------------------------------
