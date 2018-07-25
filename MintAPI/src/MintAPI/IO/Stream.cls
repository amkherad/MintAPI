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
'@PROJECT_LICENSE

''<class name="Stream" interface tags="">
''  <summary>
''    Some class description.
''  </summary>
''  <using>Some class alias.</using>
''  <idea>Me, Streaming system idea totally by Kelly Ethridge(VBCorLib)</idea>
''  <sample></sample>
''</class>

Option Base 0
Option Explicit

'=============================================
'=============================================
'=============================================
'<section Meta Data>
'
Const CLASSID As String = "Stream" '#DO NOT REMOVE IT
Const CLASSALIAS As String = APPLICATIONDOMAIN & "." & CLASSID '#DO NOT REMOVE IT
Const CLASSVERSIONSTRING As String = "0.1" '#DO NOT REMOVE IT
Const CLASSALIASSTRING As String = CLASSALIAS & "/" & CLASSVERSIONSTRING '#DO NOT REMOVE IT

'
'</section>
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section API/Type Declarations>
'
Public Enum SeekOrigin
    FromBeginning = FILE_BEGIN
    FromCurrent = FILE_CURRENT
    FromEnd = FILE_END
End Enum
Public Enum StreamReliability
    srConstant
    srVolatile
    srTimed
    srCustom
End Enum
'
'</section>
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section Interface Body Members>
'
'*********************************************
'*********************************************

Public Property Get CanTimeout() As Boolean: End Property
Public Property Get ReadTimeout() As Long: End Property
Public Property Let ReadTimeout(ByVal Value As Long): End Property
Public Property Get WriteTimeout() As Long: End Property
Public Property Let WriteTimeout(ByVal Value As Long): End Property
Public Property Get Direction() As StreamDirection: End Property
Public Property Get CanRead() As Boolean: End Property
Public Property Get CanSeek() As Boolean: End Property
Public Property Get CanWrite() As Boolean: End Property
Public Property Get Length() As Long: End Property
Public Property Get Length64() As Int64: End Property
Public Property Get Position() As Long: End Property
Public Property Let Position(ByVal Value As Long): End Property
Public Property Get Position64() As Int64: End Property
Public Property Let Position64(ByRef Value As Int64): End Property
Public Function BeginRead(ByRef Buffer() As Byte, ByVal Offset As Long, ByVal Count As Long, Optional ByVal Callback As AsyncCallback, Optional ByVal State As Variant) As IAsyncResult: End Function
Public Function BeginWrite(ByRef Buffer() As Byte, ByVal Offset As Long, ByVal Count As Long, Optional ByVal Callback As AsyncCallback, Optional ByVal State As Variant) As IAsyncResult: End Function
Public Function EndRead(ByVal AsyncResult As IAsyncResult) As Long: End Function
Public Sub EndWrite(ByVal AsyncResult As IAsyncResult): End Sub
Public Sub Flush(): End Sub
Public Function ReadBlock(ByRef Buffer() As Byte, ByVal Offset As Long, ByVal Count As Long) As Long: End Function
Public Function ReadByte() As Long: End Function
Public Function SeekPosition(ByVal Offset As Currency, ByVal Origin As SeekOrigin) As Int64: End Function
Public Sub SetLength(Value As Int64): End Sub
Public Sub WriteBlock(ByRef Buffer() As Byte, ByVal Offset As Long, ByVal Count As Long): End Sub
Public Sub WriteByte(ByVal Value As Byte): End Sub
Public Sub CloseStream(): End Sub
Public Function ToString() As String: End Function
Public Function Equals(ByRef Value As Variant) As Boolean: End Function
Public Function GetHashCode() As Long: End Function

'*********************************************
'*********************************************
'
'</section>
'---------------------------------------------
'---------------------------------------------

