VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ApplicationStream"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'@PROJECT_LICENSE

''<class name="ApplicationStream" tags="">
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
Implements IClassStream


'=============================================
'=============================================
'=============================================
'<section Meta Data>
'
Const CLASSID As String = "ApplicationStream" '#DO NOT REMOVE IT
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

Dim p_Direction  As StreamDirection
Dim p_ReadState As Boolean
Dim p_WriteState As Boolean
Dim p_Hndl As Long
Dim p_Name As String

Dim inState As Boolean, outState As Boolean

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
'Friend Sub Constructor(ByVal CInfo As ObjectConstructionInfo)
'
'    This_Constructed = True
'End Sub
Friend Sub Constructor0(ByVal Handle As Long, ByVal Direction As StreamDirection, ByVal DirName As String)
    p_Hndl = Handle
    p_ReadState = (Direction And sdInStream)
    p_WriteState = (Direction And sdOutStream)
    p_Name = DirName
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
    Call API_CloseHandle(p_Hndl)
    p_Hndl = 0
End Sub

Private Property Get Stream_Direction() As StreamDirection
    Stream_Direction = sdBoth
End Property

Private Function Stream_EndRead(ByVal AsyncResult As IAsyncResult) As Long
    
End Function
Private Sub Stream_EndWrite(ByVal AsyncResult As IAsyncResult)
    
End Sub

Private Function Stream_Equals(Value As Variant) As Boolean
    If IsObject(Value) Then
        Stream_Equals = (Value Is Me)
    End If
End Function

''<summary>
''  The function does not return until the client has
''  read all buffered data from the pipe.
''</summary>
Private Sub Stream_Flush()
    If Not p_WriteState Then throw Exps.InvalidOperationException
    If API_FlushFileBuffers(p_Hndl) = NO_VALUE Then _
        throw Exps.IfError
End Sub

Private Function Stream_GetHashCode() As Long
    Stream_GetHashCode = ObjPtr(CUnk(Me))
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
    If Not Arrays.EnsureArrayRange1DSAPtr(Arrays.GetSafeArrayPointer(Buffer), Offset, Count, outExp) Then _
        throw outExp
    
    If Not p_ReadState Then _
        throw Exps.InvalidOperationException("The stream does not support reading.")
    
    Dim outReadLen As Long
    
    If API_ReadFile(p_Hndl, Buffer(Offset), Count, outReadLen, ByVal 0&) = NO_VALUE Then _
        throw Exps.IfError
    Stream_ReadBlock = outReadLen
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
Private Function Stream_SeekPosition(ByVal Offset As Currency, ByVal Origin As SeekOrigin) As Int64
    throw Exps.NotSupportedException
End Function
Private Sub Stream_SetLength(Value As Int64)
    throw Exps.NotSupportedException
End Sub

Private Function Stream_ToString() As String
    Stream_ToString = CLASSALIASSTRING
End Function

Private Sub Stream_WriteBlock(Buffer() As Byte, ByVal Offset As Long, ByVal Count As Long)
    Dim outExp As Exception
    If Not Arrays.EnsureArrayRange1DSAPtr(Arrays.GetSafeArrayPointer(Buffer), Offset, Count, outExp) Then _
        throw outExp
    
    If Not p_WriteState Then _
        throw Exps.InvalidOperationException("The stream does not support writing.")
    
    If API_WriteFile(p_Hndl, Buffer(Offset), Count, 0, ByVal 0&) = NO_VALUE Then _
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
'<section IClassStream Implementation>
'
Private Function IClassStream_Available(ByVal Direction As StreamDirection) As Long
    
End Function
Private Sub IClassStream_CloseStream(Optional ByVal Direction As StreamDirection = 3&)
    If (Direction And sdInStream) = sdInStream Then inState = False
    If (Direction And sdOutStream) = sdOutStream Then outState = False
    If Not (p_ReadState = ((Direction And sdInStream) = sdInStream) Or _
            p_WriteState = ((Direction And sdOutStream) = sdOutStream)) Then _
       throw Exps.InvalidOperationException
    Call Stream_CloseStream
End Sub
Private Function IClassStream_GetState(ByVal Dir As StreamDirection) As Boolean
    Dim RetVal As Boolean
    If (Dir And sdInStream) = sdInStream Then
        RetVal = p_ReadState And inState
    End If
    If (Dir And sdOutStream) = sdOutStream Then
        RetVal = RetVal Or (p_WriteState And outState)
    End If
    IClassStream_GetState = RetVal
End Function
Private Sub IClassStream_InStream(ByVal Buffer As ClassStreamBuffer, Optional ByVal Length As Long = -1&)
    If Buffer Is Nothing Then throw Exps.ArgumentNullException("Buffer")
    Select Case (Buffer.BufferType And csbtTypeMask)
        Case ClassStreamBufferType.csbtNewLine
            Dim BANL As ByteArray
            Set BANL = ByteArray(vbNewLine)
            Call Stream_WriteBlock(BANL.ToConstData, 0, BANL.Length)
        Case ClassStreamBufferType.csbtFlush
            Call Stream_Flush
        Case ClassStreamBufferType.csbtClear
            throw Exps.NotSupportedException
        Case Else
            Dim bBuffer(1000) As Byte, Readed As Long
            Readed = Buffer.ReadLength
            If Readed <= 0 Then Readed = 100
            Readed = Stream_ReadBlock(bBuffer, 0, Readed)
            Call Buffer.SetBuffer(bBuffer, Readed)
    End Select
End Sub
Private Sub IClassStream_OpenStream(Optional ByVal Direction As StreamDirection = 3&)
    If (Direction And sdInStream) = sdInStream Then inState = True
    If (Direction And sdOutStream) = sdOutStream Then outState = True
End Sub
Private Sub IClassStream_OutStream(ByVal Buffer As ClassStreamBuffer, Optional ByVal Length As Long = -1&)
    If Buffer Is Nothing Then throw Exps.ArgumentNullException("Buffer")
    Select Case (Buffer.BufferType And csbtTypeMask)
        Case ClassStreamBufferType.csbtNewLine
            Dim BANL As ByteArray
            Set BANL = ByteArray(vbNewLine)
            Call Stream_WriteBlock(BANL.ToConstData, 0, BANL.Length)
        Case ClassStreamBufferType.csbtFlush
            Call Stream_Flush
        Case ClassStreamBufferType.csbtClear
            throw Exps.NotSupportedException
        Case Else
            Dim BA As ByteArray
            Set BA = Buffer.Buffer
            Call Stream_WriteBlock(BA.ToConstData, 0, BA.Length)
    End Select
End Sub
Private Function IClassStream_ReadBytes(ByVal Length As Long, Optional ByVal StartAt As Long = 0&) As Byte()
    
End Function
Private Sub IClassStream_SetInState(Optional ByVal State As Boolean = True, Optional Reserved As Variant)
    Call mint_setstream_state(State, inState, "", Reserved)
End Sub
Private Sub IClassStream_SetOutState(Optional ByVal State As Boolean = True, Optional Reserved As Variant)
    Call mint_setstream_state(State, outState, "", Reserved)
End Sub
'
'</section>
'---------------------------------------------
'---------------------------------------------
