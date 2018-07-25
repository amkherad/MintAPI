VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ByteArray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements ITargetStream
Implements IClassTexer

Dim d() As Byte
Dim dlen As Long
Dim inState As Boolean
Dim outState As Boolean

Friend Sub Initialize(Optional target)
    If Not IsMissing(target) Then Call setBuffer(target)
End Sub

Private Sub Class_Initialize()
    inState = False
    outState = False
End Sub

Public Property Let Length(Value As Long)
    Call reserve(Value, True)
End Property
Public Property Get Length() As Long
    Length = dlen
End Property
Public Sub Clear()
    Erase d
    dlen = 0
End Sub
Public Sub reserve(Length As Long, Optional keepLoadedData As Boolean = False)
    If Length >= 0 Then throw InvalidArrayLBoundException("Length Can't Be Zero Or Negative.")
    If keepLoadedData Then
        ReDim Preserve d(Length - 1)
    Else
        ReDim d(Length - 1)
    End If
    dlen = Length
End Sub
Public Sub setBuffer(targetData)
    Dim bArray() As Byte
    Call API_CopyMemory(ByVal API_VarPtrArray(bArray), ByVal API_VarPtrArray(mint_get_byte_array_of(targetData)), 4)
    'gets all types of data into byte array.
    Dim tdlen As Long
    tdlen = ArraySize(bArray)
    If dlen > tdlen Then Call Clear
    If dlen < tdlen Then Call reserve(tdlen)
    
    Dim i As Long
    For i = 0 To tdlen - 1
        d(i) = bArray(i)
    Next
End Sub
Public Sub setBoundedBuffer(targetByteArray() As Byte)
    Dim blen As Long
    blen = ArraySize(targetByteArray)
    If blen <= 0 Then throw InvalidArgumentValueException("targetByteArray Is Null.")
    If LBound(targetByteArray) <> 0 Then throw InvalidArrayLBoundException("Invalid targetByteArray LBound.")
    Erase d
    dlen = 0
    Call API_CopyMemory(ByVal API_VarPtrArray(d), ByVal API_VarPtrArray(targetByteArray), 4)
    dlen = blen
End Sub
    
Public Sub Append(target)
    Dim targetArray() As Byte
    targetArray = mint_get_byte_array_of(target)
    If ArraySize(targetArray) = 0 Then Exit Sub
    Call AppendArrayToArray(d, targetArray)
    dlen = ArraySize(d)
End Sub
Public Sub CopyByteArray(targetByteArray() As Byte)
    targetByteArray = d
End Sub
Public Function GetByteArray() As Byte()
    Call API_CopyMemory(ByVal API_VarPtrArray(GetByteArray), ByVal API_VarPtrArray(d), 4)
End Function
Public Sub SetByteArray(b() As Byte)
    Dim tdlen As Long
    tdlen = ArraySize(b)
    If dlen > tdlen Then Call Clear
    If dlen < tdlen Then Call reserve(tdlen)
    
    Dim i As Long
    For i = 0 To tdlen - 1
        d(i) = b(i)
    Next
End Sub
Public Function toByteArray() As Byte()
    toByteArray = d
End Function

Public Function toInt() As Integer
    If dlen = 0 Then Exit Function
    If dlen <= 2 Then
        toInt = ByteArrayToInteger(d)
    Else
        throw OutOfRangeException
    End If
End Function
Public Function toString() As String
    If dlen = 0 Then Exit Function
    toString = ByteArrayToString(d)
End Function
Public Function toLong() As Long
    If dlen = 0 Then Exit Function
    If dlen <= 4 Then
        toLong = ByteArrayToLong(d)
    Else
        throw OutOfRangeException
    End If
End Function
Public Function toDouble() As Double
    If dlen = 0 Then Exit Function
    If dlen <= 8 Then
        toLong = ByteArrayToDouble(d)
    Else
        throw OutOfRangeException
    End If
End Function
Public Function toFloat() As Single
    If dlen = 0 Then Exit Function
    If dlen <= 4 Then
        toLong = ByteArrayToSingle(d)
    Else
        throw OutOfRangeException
    End If
End Function
Public Function toSingle() As Single
    If dlen = 0 Then Exit Function
    If dlen <= 4 Then
        toLong = ByteArrayToSingle(d)
    Else
        throw OutOfRangeException
    End If
End Function
Public Function toDate() As Date
    If dlen = 0 Then Exit Function
    If dlen <= 8 Then
        toLong = ByteArrayToDate(d)
    Else
        throw OutOfRangeException
    End If
End Function
Public Function toDateTime() As DateTime
    If Not checkifspc(d, "datetime") Then GoTo throwInvalid
    
    Dim dt As New DateTime
    Call dt.Initialize(Now)
    Call dt.FromByteArrayData(GetByteArrayInnerData(d))
    Set toDateTime = dt
    
    Exit Function
throwInvalid:
    throw InvalidCallException("Cant convert bytearray to datetime.")
End Function
Public Function toDirectory() As Directory
    Dim d As New Directory
    Call d.Initialize(toString())
    Set toDirectory = d
End Function
Public Function toFile() As File
    Dim f As New File
    Call f.Initialize(toString())
    Set toFile = f
End Function
Public Function toEventArgs() As IEventArgs
    
End Function
Public Function toApplication() As IApplication
    
End Function
Public Function toNoteBuffer() As NoteBuffer
    
End Function
Public Function toStringParser() As StringParser
    Dim sp As New StringParser
    Call sp.Initialize(toString())
    Set toStringParser = sp
End Function
Public Function toLanguage() As Language
    
End Function
Public Function toPixmap() As Pixmap
    If Not checkifspc(d, "pixmap") Then GoTo throwInvalid
    
    Dim pm As New Pixmap
    
    Set toPixmap = pm
    
    Exit Function
throwInvalid:
    throw InvalidCallException("Cant convert bytearray to pixmap.")
End Function
Public Function toData() As IData
    
End Function
Public Function toException() As Exception
    
End Function
Public Function toAction() As Action
    Dim act As New Action
    Call act.Initialize(toString())
    Set toAction = act
End Function
Public Function toConfiguration() As Configuration
    
End Function


Public Sub fromInt(Value As Integer)
    
End Sub
Public Sub fromString(Value As String)
    
End Sub
Public Sub fromLong(Value As Long)
    
End Sub
Public Sub fromDouble(Value As Double)
    
End Sub
Public Sub fromFloat(Value As Single)
    
End Sub
Public Sub fromSingle(Value As Single)
    
End Sub
Public Sub fromDate(Value As Date)
    
End Sub
Public Sub fromDateTime(Value As DateTime)
    
End Sub
Public Sub fromDirectory(Value As Directory)
    
End Sub
Public Sub fromFile(Value As File)
    
End Sub
Public Sub fromEventArgs(Value As EventArgs)
    
End Sub
Public Sub fromApplication(Value As IApplication)
    
End Sub
Public Sub fromNoteBuffer(Value As NoteBuffer)
    
End Sub
Public Sub fromStringParser(Value As StringParser)
    
End Sub
Public Sub fromLanguage(Value As Language)
    
End Sub
Public Sub fromPixmap(Value As Pixmap)
    
End Sub
Public Sub fromData(Value As IData)
    
End Sub
Public Sub fromException(Value As Exception)
    
End Sub
Public Sub fromAction(Value As Action)
    
End Sub
Public Sub fromConfiguration(Value As Configuration)
    
End Sub

'$ByteArray$08datetime3###
Private Function checkifspc(b() As Byte, className As String) As Boolean
    If Not IsByteArraySpecifiedData(b) Then
        checkifspc = False
        Exit Function
    End If
    If GetByteArrayClassTypeName(b) <> className Then
        checkifspc = False
        Exit Function
    End If
End Function
Private Sub IsByteArraySpecifiedData(b() As Byte)
On Error GoTo err
    Dim header() As Byte
    header = GetSubArray(b, 0, 11)
    
    If ArrayCompare(header, StringToByteArray("$ByteArray$")) Then
        IsByteArraySpecifiedData = True
    Else
        IsByteArraySpecifiedData = False
    End If
err:
End Sub
Private Function GetByteArrayClassTypeName(b() As Byte) As String
On Error GoTo returnNA
    If Not IsByteArraySpecifiedData(b) Then GoTo returnNA
    Dim bSize As Long
    bSize = ArraySize(b)
    If bSize <= 0 Then GoTo returnNA
    
    Dim ln() As Byte, afterCLSNameLength As Integer
    ln = GetSubArray(b, 11, 2)
    
    afterCLSNameLength = ByteArray(b).toInt
    If afterCLSNameLength <= 0 Then GoTo returnNA
    
    ln = GetSubArray(b, 13, afterCLSNameLength)
    If ArraySize(ln) <= 0 Then GoTo returnNA
    
    GetByteArrayClassTypeName = ByteArrayToString(ln)
    
    Exit Function
returnNA:
    GetByteArrayClassTypeName = "N/A"
End Function
Private Function GetByteArrayInnerData(b() As Byte) As Byte()
    On Error GoTo returnN
    Dim bSize As Long
    bSize = ArraySize(b)
    If bSize <= 0 Then GoTo returnN
    
    Dim retVal() As Byte
    
    Dim ln() As Byte, afterCLSNameLength As Integer, innerDataLength As Integer
    ln = GetSubArray(b, 11, 2)
    
    afterCLSNameLength = ByteArray(ln).toInt
    If afterCLSNameLength <= 0 Then GoTo returnN
    
    ln = GetSubArray(b, afterCLSNameLength + 11, 2)
    innerDataLength = ByteArray(ln).toInt
    
    GetByteArrayInnerData = GetSubArray(b, GetByteArrayInnerData + 11, innerDataLength)
    
    Exit Function
returnN:
End Function


Private Sub IData_Append(target As Variant): Call Append(target): End Sub

Private Sub IData_Clear(): Call Clear: End Sub
Private Sub IData_CopyByteArray(targetByteArray() As Byte): Call CopyByteArray(targetByteArray): End Sub

Private Function IData_GetByteArray() As Byte()
    Call API_CopyMemory(ByVal API_VarPtrArray(IData_GetByteArray), ByVal API_VarPtrArray(d), 4)
End Function
Private Property Get IData_Length() As Long: IData_Length = dlen: End Property
Private Sub IData_SetByteArray(b() As Byte)
    Call SetByteArray(b)
End Sub
Private Function IData_toBinary() As BinaryData
    Dim binary As New BinaryData
    Call binary.setBuffer(d)
    Set IData_toBinary = BinaryData
End Function
Private Function ITargetStream_Available(dir As StreamDirection) As Long
    Dim val As Boolean
    If (dir And sdInStream) = sdInStream Then
        val = ((dlen > 0) And inState)
    End If
    If (dir And sdOutStream) = sdOutStream Then
        val = val Or (outState)
    End If
    ITargetStream_Available = val
End Function

Private Sub ITargetStream_CloseStream()
    inState = False
    outState = False
End Sub
Private Sub ITargetStream_OpenStream(Optional direction As StreamDirection = StreamDirection.sdBoth)
    If (direction And sdInStream) = sdInStream Then
        inState = True
    End If
    If (direction And sdOutStream) = sdOutStream Then
        outState = True
    End If
End Sub

Private Function ITargetStream_getState(dir As StreamDirection) As Boolean
    Dim val As Boolean
    If (dir And sdInStream) = sdInStream Then
        val = inState
    End If
    If (dir And sdOutStream) = sdOutStream Then
        val = val Or outState
    End If
    ITargetStream_getState = val
End Function

Private Sub ITargetStream_inStream(Data, Optional Length As Long = -1)
    Call mint_put_byte_array_to(Data, Length)
End Sub
Private Sub ITargetStream_outStream(Data, Optional Length As Long = -1)
    Call Append(mint_get_byte_array_of(Data, Length))
End Sub

Private Sub ITargetStream_setInState(Optional ByVal State As Boolean = True, Optional Reserved): inState = State: End Sub
Private Sub ITargetStream_setOutState(Optional ByVal State As Boolean = True, Optional Reserved): outState = State: End Sub

