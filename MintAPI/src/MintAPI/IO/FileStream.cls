VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FileStream"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"MetaObject"
'@PROJECT_LICENSE

''<class name="FileStream" tags="">
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
Implements Stream

'=============================================
'=============================================
'=============================================
'<section Meta Data>
'
Const CLASSID As String = "FileStream" '#DO NOT REMOVE IT
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
Dim myMetaObject As MetaObject

Dim p_Hndl As Long
Dim p_ReadState As Boolean
Dim p_WriteState As Boolean
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
Friend Sub Constructor0(ByVal FileHandle As Long)
    p_Hndl = FileHandle
    
    This_Constructed = True
End Sub
Friend Sub Constructor1(ByVal Path As String)
    
    
    This_Constructed = True
End Sub
Friend Sub ConstructorCopy(ByVal FileStream As FileStream)
    p_Hndl = FileStream.Handle
    p_ReadState = FileStream.CanRead
    p_WriteState = FileStream.CanWrite
    
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


Public Property Get Handle() As Long
THROW_IfNull
    Handle = p_Hndl
End Property


Public Property Get CanRead() As Boolean
THROW_IfNull
    
End Property
Public Property Get CanWrite() As Boolean
THROW_IfNull
    
End Property
Public Property Get Direction() As StreamDirection
THROW_IfNull
    
End Property
Public Property Get CanSeek() As Boolean
THROW_IfNull
    CanSeek = True
End Property
Public Property Get CanTimeout() As Boolean
THROW_IfNull
    CanTimeout = True
End Property

Public Property Get Length() As Long
THROW_IfNull
    
End Property
Public Property Get Length64() As Int64
THROW_IfNull
    Length64 = Convert.CInt64(Length)
End Property

Public Property Let Position(ByVal RHS As Long)
THROW_IfNull
    
End Property
Public Property Get Position() As Long
THROW_IfNull
    
End Property

Public Property Let Position64(RHS As Int64)
THROW_IfNull
    Position = Convert.ToLong(RHS)
End Property
Public Property Get Position64() As Int64
THROW_IfNull
    Position64 = Convert.ToInt64(Position)
End Property

Public Property Let ReadTimeout(ByVal RHS As Long)
THROW_IfNull
    
End Property
Public Property Get ReadTimeout() As Long
THROW_IfNull
    
End Property

Public Property Get WriteTimeout() As Long
THROW_IfNull
    
End Property
Public Property Let WriteTimeout(ByVal RHS As Long)
THROW_IfNull
    
End Property

Public Function SeekPosition(ByVal Offset As Currency, ByVal Origin As SeekOrigin) As Int64
THROW_IfNull
    
End Function

Public Sub SetLength(Value As Int64)
THROW_IfNull
    
End Sub

Public Sub CloseStream()
    If p_Hndl <> NO_VALUE Then
        p_WriteState = False
        p_ReadState = False
        Call API_CloseHandle(p_Hndl)
        p_Hndl = NO_VALUE
    End If
End Sub

Public Sub Flush()
THROW_IfNull
    If Not p_WriteState Then throw Exps.IOException
    If API_FlushFileBuffers(p_Hndl) = NO_VALUE Then _
        throw Exps.IfError
End Sub

Public Sub WriteBlock(Buffer() As Byte, ByVal Offset As Long, ByVal Count As Long)
THROW_IfNull
    Dim outExp As Exception
    If Not Arrays.EnsureArrayRange1DSAPtr(Arrays.GetSafeArrayPointer(Buffer), Offset, Count, outExp) Then _
        throw outExp
    
    If Not p_WriteState Then _
        throw Exps.IOException("The stream does not support writing.")
    
    If API_WriteFile(p_Hndl, Buffer(Offset), Count, 0, ByVal 0&) = NO_VALUE Then _
        throw Exps.IfError
End Sub
Public Function ReadBlock(Buffer() As Byte, ByVal Offset As Long, ByVal Count As Long)
THROW_IfNull
    Dim outExp As Exception
    If Not Arrays.EnsureArrayRange1DSAPtr(Arrays.GetSafeArrayPointer(Buffer), Offset, Count, outExp) Then _
        throw outExp
    
    If Not p_ReadState Then _
        throw Exps.IOException("The stream does not support writing.")
    
    Dim outReadLen As Long
    
    If API_ReadFile(p_Hndl, Buffer(Offset), Count, outReadLen, ByVal 0&) = NO_VALUE Then _
        throw Exps.IfError
    ReadBlock = outReadLen
End Function

Public Sub WriteByte(ByVal Value As Byte)
THROW_IfNull
    Dim B(0) As Byte
    B(0) = Value
    Call WriteBlock(B, 0, 1)
End Sub
Public Function ReadByte() As Long
THROW_IfNull
    Dim B(0) As Byte
    Call ReadBlock(B, 0, 1)
    ReadByte = B(0)
End Function

Public Function BeginRead(Buffer() As Byte, ByVal Offset As Long, ByVal Count As Long, Optional ByVal Callback As AsyncCallback, Optional ByVal State As Variant) As IAsyncResult
THROW_IfNull
    
End Function

Public Function BeginWrite(Buffer() As Byte, ByVal Offset As Long, ByVal Count As Long, Optional ByVal Callback As AsyncCallback, Optional ByVal State As Variant) As IAsyncResult
THROW_IfNull
    
End Function

Public Function EndRead(ByVal AsyncResult As IAsyncResult) As Long
THROW_IfNull
    
End Function
Public Sub EndWrite(ByVal AsyncResult As IAsyncResult)
THROW_IfNull
    
End Sub

'=============================================
'=============================================
'=============================================
'<section Stream Implementation>
'
Private Property Get Stream_CanRead() As Boolean: Stream_CanRead = CanRead: End Property
Private Property Get Stream_CanSeek() As Boolean: Stream_CanSeek = CanSeek: End Property
Private Property Get Stream_CanTimeout() As Boolean: Stream_CanTimeout = CanTimeout: End Property
Private Property Get Stream_CanWrite() As Boolean: Stream_CanWrite = CanWrite: End Property
Private Property Get Stream_Direction() As StreamDirection: Stream_Direction = Direction: End Property
Private Property Get Stream_WriteTimeout() As Long: Stream_WriteTimeout = WriteTimeout: End Property
Private Sub Stream_CloseStream(): Call CloseStream: End Sub
Private Function Stream_BeginRead(Buffer() As Byte, ByVal Offset As Long, ByVal Count As Long, Optional ByVal Callback As AsyncCallback, Optional ByVal State As Variant) As IAsyncResult: Set Stream_BeginRead = BeginRead(Buffer, Offset, Count, Callback, State): End Function
Private Function Stream_BeginWrite(Buffer() As Byte, ByVal Offset As Long, ByVal Count As Long, Optional ByVal Callback As AsyncCallback, Optional ByVal State As Variant) As IAsyncResult: Set Stream_BeginWrite = BeginWrite(Buffer, Offset, Count, Callback, State): End Function
Private Function Stream_EndRead(ByVal AsyncResult As IAsyncResult) As Long: Stream_EndRead = EndRead(AsyncResult): End Function
Private Sub Stream_EndWrite(ByVal AsyncResult As IAsyncResult): Call EndWrite(AsyncResult): End Sub
Private Sub Stream_Flush(): Call Flush: End Sub
Private Property Get Stream_Length() As Long: Stream_Length = Length: End Property
Private Property Get Stream_Length64() As Int64: Stream_Length64 = Length64: End Property
Private Property Get Stream_Position() As Long: Stream_Position = Position: End Property
Private Property Let Stream_Position(ByVal RHS As Long): Position = RHS: End Property
Private Property Get Stream_Position64() As Int64: Stream_Position64 = Position64: End Property
Private Property Let Stream_Position64(RHS As Int64): Position64 = RHS: End Property
Private Function Stream_ReadBlock(Buffer() As Byte, ByVal Offset As Long, ByVal Count As Long) As Long: Stream_ReadBlock = Stream_ReadBlock(Buffer, Offset, Count): End Function
Private Function Stream_ReadByte() As Long: Stream_ReadByte = ReadByte: End Function
Private Property Let Stream_ReadTimeout(ByVal RHS As Long): ReadTimeout = RHS: End Property
Private Property Get Stream_ReadTimeout() As Long: Stream_ReadTimeout = ReadTimeout: End Property
Private Function Stream_SeekPosition(ByVal Offset As Currency, ByVal Origin As SeekOrigin) As Int64: Stream_SeekPosition = SeekPosition(Offset, Origin): End Function
Private Sub Stream_SetLength(Value As Int64): Call SetLength(Value): End Sub
Private Sub Stream_WriteBlock(Buffer() As Byte, ByVal Offset As Long, ByVal Count As Long): Call WriteBlock(Buffer, Offset, Count): End Sub
Private Sub Stream_WriteByte(ByVal Value As Byte): Call WriteByte(Value): End Sub
Private Property Let Stream_WriteTimeout(ByVal RHS As Long): WriteTimeout = RHS: End Property

Private Function Stream_ToString() As String: Stream_ToString = ToString: End Function
Private Function Stream_GetHashCode() As Long: Stream_GetHashCode = GetHashCode: End Function
Private Function Stream_Equals(Value As Variant) As Boolean: Stream_Equals = Equals(Value): End Function
'
'</section>
'---------------------------------------------
'---------------------------------------------

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

Public Function Clone() As FileStream
    Set Clone = New FileStream
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
Private Sub Class_Initialize()
    
End Sub
Private Sub Class_Terminate()
    If Not myMetaObject Is Nothing Then
        Signal.Emit Me, SIGNAL_TERMINATING
        Call myMetaObject.Dispose
        Set myMetaObject = Nothing
    End If
End Sub

Private Sub THROW_IfNull()
    If p_Hndl = NO_VALUE Then throw Exps.InvalidOperationException
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
