VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ProcessStream"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'@PROJECT_LICENSE

''<class name="ApplicationStandardIO" tags="">
''  <summary>
''    Some class description.
''  </summary>
''  <constructors>
''      <constructor name="" friend>
''          Some constructor description.
''      </constructor>
''  </constructors>
''  <using>Stream</using>
''  <idea>Me</idea>
''  <sample></sample>
''</class>

Option Base 0
Option Explicit

Implements Stream


'=============================================
'=============================================
'=============================================
'<section Meta Data>
'
Const CLASSID As String = "ApplicationStandardIO" '#DO NOT REMOVE IT
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
Dim This_Constructed As Boolean
'Dim myMetaObject As MetaObject

Dim p_direction  As StreamDirection
Dim p_ReadState As Boolean
Dim p_WriteState As Boolean
Dim p_hndl As Long
Dim p_name As String

Dim p_AsyncThread As Thread
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
'Friend Sub ConstructorArg(ByVal Arguments As ArgumentList)
'
'    This_Constructed = True
'End Sub
Friend Sub Constructor0(ByVal Handle As Long, ByVal Direction As StreamDirection)
    p_hndl = Handle
    
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


Private Function Stream_BeginRead(Buffer() As Byte, ByVal Offset As Long, ByVal Count As Long, Optional ByVal Callback As AsyncCallback, Optional ByVal State As Variant) As IAsyncResult: Async
    
End Function
Private Function Stream_BeginWrite(Buffer() As Byte, ByVal Offset As Long, ByVal Count As Long, Optional ByVal Callback As AsyncCallback, Optional ByVal State As Variant) As IAsyncResult: Async
    
End Function

Private Property Get Stream_CanRead() As Boolean
    Stream_CanRead = p_ReadState
End Property

Private Property Get Stream_CanSeek() As Boolean
    Stream_CanSeek = False
End Property

Private Property Get Stream_CanTimeout() As Boolean
    Stream_CanTimeout = False
End Property

Private Property Get Stream_CanWrite() As Boolean
    Stream_CanWrite = p_WriteState
End Property

Private Sub Stream_CloseStream()
    p_WriteState = False
    p_ReadState = False
    Call API_CloseHandle(p_hndl)
    p_hndl = 0
End Sub

Private Function Stream_EndRead(ByVal AsyncResult As IAsyncResult) As Long
    
End Function
Private Sub Stream_EndWrite(ByVal AsyncResult As IAsyncResult)
    
End Sub

Private Function Stream_Equals(Value As Variant) As Boolean
    If IsObject(Value) Then
        Stream_Equals = (Value Is Me)
    End If
End Function

Private Sub Stream_Flush()
    If Not p_WriteState Then throw Exps.WriteFileException
    If API_FlushFileBuffers(p_hndl) = NO_VALUE Then _
        throw Exps.IfError
End Sub

Private Function Stream_GetHashCode() As Long
    FlushFileBuffers = ObjPtr(CUnk(Me))
End Function

Private Property Get Stream_Length() As Long
    throw Exps.NotSupportedException
End Property
Private Property Get Stream_Length64() As Int64
    throw Exps.NotSupportedException
End Property
Private Property Let Stream_Position(ByVal RHS As Long)
    throw Exps.NotSupportedException
End Property
Private Property Get Stream_Position() As Long
    throw Exps.NotSupportedException
End Property
Private Property Let Stream_Position64(ByRef RHS As Int64)
    throw Exps.NotSupportedException
End Property
Private Property Get Stream_Position64() As Int64
    throw Exps.NotSupportedException
End Property

Private Function Stream_ReadBlock(Buffer() As Byte, ByVal Offset As Long, ByVal Count As Long) As Long
    Dim outExp As Exception
    If Not EnsureArrayRange1D(GetArrayPointer(Buffer), Offset, Count, outExp) Then _
        throw outExp
    
    If Not p_WriteState Then _
        throw Exps.WriteFileException("The stream does not support writing.")
    
    If API_WriteFile(p_hndl, Buffer(Offset), Count, 0, ByVal 0&) = NO_VALUE Then _
        throw Exps.IfError
End Function
Private Function Stream_ReadByte() As Long
    Dim B(0) As Byte
    Call Stream_ReadBlock(B, 0, 1)
    Stream_ReadByte = B(0)
End Function

Private Property Let Stream_ReadTimeout(ByVal RHS As Long)
    throw Exps.NotSupportedException
End Property
Private Property Get Stream_ReadTimeout() As Long
    throw Exps.NotSupportedException
End Property
Private Function Stream_SeekPosition(ByVal Offset As Currency, ByVal Origin As SeekOrigin) As BigNumber
    throw Exps.NotSupportedException
End Function
Private Sub Stream_SetLength(ByVal Value As BigNumber)
    throw Exps.NotSupportedException
End Sub

Private Function Stream_ToString() As String
    Stream_ToString = CLASSALIASSTRING
End Function

Private Sub Stream_WriteBlock(Buffer() As Byte, ByVal Offset As Long, ByVal Count As Long)
    Dim outExp As Exception
    If Not EnsureArrayRange1D(GetArrayPointer(Buffer), Offset, Count, outExp) Then _
        throw outExp
    
    If Not p_WriteState Then _
        throw Exps.WriteFileException("The stream does not support writing.")
    
    If API_WriteFile(p_hndl, Buffer(Offset), Count, 0, ByVal 0&) = NO_VALUE Then _
        throw Exps.IfError
End Sub
Private Sub Stream_WriteByte(ByVal Value As Byte)
    Dim B(0) As Byte
    B(0) = Value
    Call Stream_WriteBlock(B, 0, 1)
End Sub

Private Property Let Stream_WriteTimeout(ByVal RHS As Long)
    throw Exps.NotSupportedException
End Property
Private Property Get Stream_WriteTimeout() As Long
    throw Exps.NotSupportedException
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
'Public Property Get Alias() As String
'    Alias = CLASSALIASSTRING
'End Property
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
'    Set MetaObject = InitMetaObject
'End Function
'
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
'
'End Sub
'Private Sub Class_Terminate()
'    Signal.Emit Me, SIGNAL_TERMINATING
'    If Not myMetaObject Is Nothing Then
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
'Private Sub IObject_Initialize(ByVal Arguments As ArgumentList)
'    Call ConstructorArg(Arguments)
'End Sub
'Private Property Get IObject_Alias() As String
'    IObject_Alias = CLASSALIASSTRING
'End Property
'Private Function IObject_Equals(Value As Variant) As Boolean
'    IObject_Equals = Equals(Value)
'End Function
'Private Function IObject_GetHashCode() As Long
'    IObject_GetHashCode = GetHashCode
'End Function
'Private Function IObject_MetaObject() As MetaObject
'    Set IObject_MetaObject = MetaObject
'End Function
'Private Function IObject_ToString() As String
'    IObject_ToString = ToString
'End Function
'Private Function IObject_Notify(ByVal Name As String, ByVal Arguments As ArgumentList) As Variant
'    Select Case LCase$(Name)
'        'Case ""
'
'        Case Else
'            throw UnknownNotificationException, Me
'    End Select
'End Function
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
'<section MetaObject constructor>
'
'Private Function InitMetaObject() As MetaObject
'    If Not myMetaObject Is Nothing Then
'        Set InitMetaObject = myMetaObject
'    End If
'    Call Objects.MetaObject.Setup(MetaObject, Me, True, False, CLASSVERSIONSTRING)
'    '...
'End Function
'
'</section>
'---------------------------------------------
'---------------------------------------------
