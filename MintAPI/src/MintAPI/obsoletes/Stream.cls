VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Stream"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"

Private Const FILE_BEGIN    As Long = 0
Private Const FILE_END      As Long = 2
Private Const FILE_CURRENT  As Long = 1

''
' The reference point in which to begin seeking from.
'
' @param FromBeginning The new position will be offset from the beginning of the file.
' A negative value will cause an exception.
' @param FromCurrent The position will be offset from the current file pointer location.
' The value can be negative to seek backwards from the current position.
' @param Fromend The new position will be offset from the end of the file. A positive
' number will cause an exception. The offset should be 0 or negative.
'
Public Enum SeekOrigin
    FromBeginning = FILE_BEGIN
    FromCurrent = FILE_CURRENT
    FromEnd = FILE_END
End Enum


Public Property Get CanTimeout() As Boolean: End Property
Public Property Get ReadTimeout() As Long: End Property
Public Property Let ReadTimeout(ByVal Value As Long): End Property
Public Property Get WriteTimeout() As Long: End Property
Public Property Let WriteTimeout(ByVal Value As Long): End Property
Public Property Get CanRead() As Boolean: End Property
Public Property Get CanSeek() As Boolean: End Property
Public Property Get CanWrite() As Boolean: End Property
Public Property Get Length() As Long: End Property
Public Property Get Length64() As Int64: End Property
Public Property Get Position() As Long: End Property
Public Property Let Position(ByVal Value As Long): End Property
Public Property Get Position64() As Int64: End Property
Public Property Let Position64(ByVal Value As Int64): End Property
Public Function BeginRead(ByRef Buffer() As Byte, ByVal Offset As Long, ByVal Count As Long, Optional ByVal Callback As AsyncCallback, Optional ByVal State As Variant) As IAsyncResult: End Function
Public Function BeginWrite(ByRef Buffer() As Byte, ByVal Offset As Long, ByVal Count As Long, Optional ByVal Callback As AsyncCallback, Optional ByVal State As Variant) As IAsyncResult: End Function
Public Sub CloseStream(): End Sub
Public Function EndRead(ByVal AsyncResult As IAsyncResult) As Long: End Function
Public Sub EndWrite(ByVal AsyncResult As IAsyncResult): End Sub
Public Function Equals(ByRef Value As Variant) As Boolean: End Function
Public Sub Flush(): End Sub
Public Function GetHashCode() As Long: End Function
Public Function ReadBlock(ByRef Buffer() As Byte, ByVal Offset As Long, ByVal Count As Long) As Long: End Function
Public Function ReadByte() As Long: End Function
Public Function SeekPosition(ByVal Offset As Currency, ByVal Origin As SeekOrigin) As BigNumber: End Function
Public Sub SetLength(ByVal Value As BigNumber): End Sub
Public Function ToString() As String: End Function
Public Sub WriteBlock(ByRef Buffer() As Byte, ByVal Offset As Long, ByVal Count As Long): End Sub
Public Sub WriteByte(ByVal Value As Byte): End Sub
