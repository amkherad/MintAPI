VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StreamAsyncResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"MetaObject"
'@PROJECT_LICENSE

''<class name="StreamAsyncResult" tags="">
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

'Implements IObject
'Implements Stream

'=============================================
'=============================================
'=============================================
'<section Meta Data>
'
Const CLASSID As String = "StreamAsyncResult" '#DO NOT REMOVE IT
Const CLASSALIAS As String = APPLICATIONDOMAIN & "." & CLASSID '#DO NOT REMOVE IT
Const CLASSVERSIONSTRING As String = "0.1" '#DO NOT REMOVE IT
Const CLASSALIASSTRING As String = CLASSALIAS & "/" & CLASSVERSIONSTRING '#DO NOT REMOVE IT

Const SIGNAL_TERMINATING As String = "terminating"
Const SLOT_METAOBJECT As String = "metaobject"
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
'Dim This_Constructed As Boolean
'Dim myMetaObject As MetaObject

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
'Friend Sub Constructor(ByVal CInfo As ObjectConstructionInfo)
'
'    This_Constructed = True
'End Sub
'Friend Sub Constructor0(ByVal ID As Long)
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
'Public Function Equals(Value As Variant) As Boolean
'    Equals = Objects.Equals(Me, Value)
'End Function
'Public Function ToString() As String
'    ToString = CLASSALIAS
'End Function
'Public Function GetHashCode() As Long
'    GetHashCode = Objects.GetHashCode(Me)
'End Function
'Public Function MetaObject() As MetaObject
'    If myMetaObject Is Nothing Then
'        Call Objects.MetaObject.Setup(myMetaObject, Me, True, CLASSALIASSTRING)
'    End If
'    Set MetaObject = myMetaObject
'End Function
'
'Public Function Clone() As StreamAsyncResult
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
'
'End Sub
'Private Sub Class_Terminate()
'    If Not myMetaObject Is Nothing Then
'        Signal.Emit Me, SIGNAL_TERMINATING
'        Call myMetaObject.Dispose
'        Set myMetaObject = Nothing
'    End If
'End Sub
'
'</section>
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section IObject Implementation>
'

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
