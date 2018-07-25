VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MCIPlayerMedia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'@PROJECT_LICENSE

''<class name="BasicPlayerMedia" tags="">
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

'=============================================
'=============================================
'=============================================
'<section Meta Data>
'
Const CLASSID As String = "BasicPlayerMedia" '#DO NOT REMOVE IT
Const CLASSALIAS As String = APPLICATIONDOMAIN & "." & CLASSID '#DO NOT REMOVE IT
Const CLASSVERSIONSTRING As String = "0.1" '#DO NOT REMOVE IT
Const CLASSALIASSTRING As String = CLASSALIAS & "/" & CLASSVERSIONSTRING '#DO NOT REMOVE IT

Const SIGNAL_TERMINATING                As String = "terminating"
Const SLOT_METAOBJECT                   As String = "metaobject"

Const SIGNAL_MEDIAOPEN                  As String = "open"
Const SIGNAL_MEDIACLOSE                 As String = "close"
Const SIGNAL_MEDIAPLAY                  As String = "play"
Const SIGNAL_MEDIASTOP                  As String = "stop"
Const SIGNAL_MEDIAPAUSE                 As String = "pause"
Const SIGNAL_MEDIARESUME                As String = "resume"
Const SIGNAL_MEDIASEEK                  As String = "seek"
Const SIGNAL_MEDIAATTRIBUTESCHANGED     As String = "attributeschanged"
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

Public Event MediaOpen(ByVal E As EventArgs)
Public Event MediaClose(ByVal E As EventArgs)
Public Event MediaPlay(ByVal E As EventArgs)
Public Event MediaStop(ByVal E As EventArgs)
Public Event MediaPause(ByVal E As EventArgs)
Public Event MediaResume(ByVal E As EventArgs)
Public Event MediaSeek(ByVal E As EventArgs)
Public Event MediaAttributesChanged(ByVal E As EventArgs)

Dim WithEvents p_Attr As MediaGenericAttributes
Attribute p_Attr.VB_VarHelpID = -1

Dim p_Path As String
Dim p_ID As Long, p_Alias As String
Dim p_ChkCheckForOpen As String
Dim p_State As MediaStreamState
Dim p_hWnd As Long
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
Friend Sub Constructor0(ByVal Path As String, ByVal ID As Long, ByVal Alias As String, ByVal hWnd As Long, ByVal IsOpen As Boolean)
    p_Path = Path
    p_ID = ID
    p_Alias = Alias
    p_ChkCheckForOpen = True
    p_State = IIf(IsOpen, mssOpen, mssNone)
    p_hWnd = hWnd
    
    This_Constructed = True
End Sub
Friend Sub Constructor1(ByVal ID As String, ByVal Alias As String, ByVal IsOpen As Boolean)
    p_ID = ID
    p_Alias = Alias
    p_ChkCheckForOpen = True
    If IsOpen Then _
        If Not IsPlaying Then _
            throw Exps.InvalidStatusException
    p_State = mssOpen
    
    This_Constructed = True
End Sub
Friend Sub Constructor2(ByVal ID As Long, ByVal Alias As String)
    p_ID = ID
    p_Alias = Alias
    p_ChkCheckForOpen = False
    p_State = mssOpen
    
    This_Constructed = True
End Sub
Friend Sub ConstructorCopy(ByVal MCIPlayerMedia As MCIPlayerMedia)
    
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

Friend Property Get ID() As Long
    ID = p_ID
End Property

Public Property Get Path() As String
    Path = p_Path
End Property

Public Property Get Alias() As String
    Alias = p_Alias
End Property

Public Property Get hWnd() As Long
    hWnd = p_hWnd
End Property

Public Property Get State() As MediaStreamState
    State = mciMode(ID)
End Property

Public Property Get IsOpen() As Boolean
    If (p_State And mssOpen) <> mssOpen Then Exit Property
    On Error GoTo CatchError
    IsOpen = (mciMode(ID) <> mssError)
    Exit Property
CatchError: IsOpen = False
End Property
Public Property Get IsPlaying() As Boolean
    If IsOpen Then Exit Property
    IsPlaying = (mciMode(ID) = mssPlaying)
End Property

Public Property Get Length() As Long
    Length = mciLength(ID)
End Property
Public Property Get Length64() As Int64
    Length64 = Convert.ToInt64(mciQueryInformationStr(Alias, "length"), True)
End Property
Public Property Get Position() As Long
    Position = mciPosition(ID)
End Property
Public Property Let Position(ByVal Value As Long)
    Call SeekMedia(FromBeginning, Value)
End Property
Public Property Get Position64() As Int64
    Position64 = Convert.ToInt64(mciQueryInformationStr(Alias, "position"), True, 0)
End Property
Public Property Let Position64(Value As Int64)
    Call SeekMedia64(FromBeginning, Value)
End Property

Public Sub SeekMedia(ByVal Origin As SeekOrigin, ByVal Offset As Long)
    
End Sub
Public Sub SeekMedia64(ByVal Origin As SeekOrigin, Offset As Int64)
    
End Sub

Public Property Get TimingFormat() As MediaTimingFormat
    
End Property

Public Sub PlayMedia(ByVal IsLoop As Boolean)
    Call mciPlayFile(ID, IsLoop, p_hWnd)
    p_State = mssPlaying
    Signal.Emit Me, SIGNAL_MEDIAPLAY
    RaiseEvent MediaPlay(EventArgs(Me))
End Sub
Public Sub StopMedia()
    Call mciExecuteFunction(ID, MCI_FUNC_STOP)
    p_State = mssStopped
    Signal.Emit Me, SIGNAL_MEDIASTOP
    RaiseEvent MediaStop(EventArgs(Me))
End Sub

Public Sub PauseMedia()
    Call mciExecuteFunction(ID, MCI_FUNC_PAUSE)
    p_State = mssPaused
    Signal.Emit Me, SIGNAL_MEDIAPAUSE
    RaiseEvent MediaPause(EventArgs(Me))
End Sub
Public Sub ResumeMedia()
    Call mciExecuteFunction(ID, MCI_FUNC_RESUME)
    p_State = mssPlaying
    Signal.Emit Me, SIGNAL_MEDIARESUME
    Signal.Emit Me, SIGNAL_MEDIAPLAY
    RaiseEvent MediaResume(EventArgs(Me))
    RaiseEvent MediaPlay(EventArgs(Me))
End Sub

Public Sub CloseMedia()
    Call mciExecuteFunction(ID, MCI_FUNC_CLOSE)
    p_State = mssClosed
    Signal.Emit Me, SIGNAL_MEDIACLOSE
    RaiseEvent MediaClose(EventArgs(Me))
End Sub

Public Property Get Attributes() As MediaGenericAttributes
    Set Attributes = p_Attr
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
    If Not myMetaObject Is Nothing Then
        Set MetaObject = myMetaObject
    End If
    Call Objects.MetaObject.Setup(myMetaObject, Me, True, CLASSALIASSTRING)
    Set MetaObject = myMetaObject
End Function

Public Function Clone() As MCIPlayerMedia
    Set Clone = New MCIPlayerMedia
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
    Set p_Attr = New MediaGenericAttributes
End Sub
Private Sub Class_Terminate()
    If Not myMetaObject Is Nothing Then
        Signal.Emit Me, SIGNAL_TERMINATING
        Call myMetaObject.Dispose
        Set myMetaObject = Nothing
    End If
    
    If IsOpen Then Call mciExecuteFunction(ID, MCI_FUNC_CLOSE)
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
            throw Exps.UnknownNotificationException
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
'<section MediaGenericAttributes Events>
'
Private Sub p_Attr_VolumeChanged(ByVal E As MintAPI.EventArgs)
    If (p_Attr.LeftBalance = 0) And (p_Attr.RightBalance = 0) Then
        Call mciExecuteFunction(ID, MCI_FUNC_SETAUDIO, "volume to " & (p_Attr.Volume * 10))
    Else
        Dim Vol As Long
        Vol = p_Attr.Volume
        Call mciExecuteFunction(ID, MCI_FUNC_SETAUDIO, "left volume to " & ((Vol * p_Attr.LeftBalance) / 100) * 10)
        Call mciExecuteFunction(ID, MCI_FUNC_SETAUDIO, "right volume to " & ((Vol * p_Attr.RightBalance) / 100) * 10)
    End If
    Signal.Emit Me, SIGNAL_MEDIAATTRIBUTESCHANGED
    RaiseEvent MediaAttributesChanged(EventArgs(Me))
End Sub
Private Sub p_Attr_BalanceChanged(ByVal E As MintAPI.EventArgs)
    If (p_Attr.LeftBalance = 0) And (p_Attr.RightBalance = 0) Then
        Call mciExecuteFunction(ID, MCI_FUNC_SETAUDIO, "volume to " & (p_Attr.Volume * 10))
    Else
        Dim Vol As Long
        Vol = p_Attr.Volume
        Call mciExecuteFunction(ID, MCI_FUNC_SETAUDIO, "left volume to " & ((Vol * p_Attr.LeftBalance) / 100) * 10)
        Call mciExecuteFunction(ID, MCI_FUNC_SETAUDIO, "right volume to " & ((Vol * p_Attr.RightBalance) / 100) * 10)
    End If
    Signal.Emit Me, SIGNAL_MEDIAATTRIBUTESCHANGED
    RaiseEvent MediaAttributesChanged(EventArgs(Me))
End Sub

Private Sub p_Attr_BassChanged(ByVal E As MintAPI.EventArgs)
    Call mciExecuteFunction(ID, MCI_FUNC_SETAUDIO, "bass to " & (p_Attr.Bass * 10))
    Signal.Emit Me, SIGNAL_MEDIAATTRIBUTESCHANGED
    RaiseEvent MediaAttributesChanged(EventArgs(Me))
End Sub
Private Sub p_Attr_TrebleChanged(ByVal E As MintAPI.EventArgs)
    Call mciExecuteFunction(ID, MCI_FUNC_SETAUDIO, "treble to " & (p_Attr.Treble * 10))
    Signal.Emit Me, SIGNAL_MEDIAATTRIBUTESCHANGED
    RaiseEvent MediaAttributesChanged(EventArgs(Me))
End Sub
'
'</section>
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section IMedia Implementations>
'
Private Property Get IMedia_Handle() As Long: IMedia_Handle = GetHashCode: End Property
Private Property Get IMedia_Stream() As Stream: End Property
Private Property Get IMedia_IsOnStorage() As Boolean: IMedia_IsOnStorage = True: End Property
Private Property Get IMedia_Path() As String: IMedia_Path = Path: End Property
Private Property Get IMedia_State() As MediaStreamState: IMedia_State = State: End Property
Private Property Get IMedia_Length() As Long: IMedia_Length = Length: End Property
Private Property Get IMedia_Length64() As Int64: IMedia_Length64 = Length64: End Property
Private Property Get IMedia_Position() As Long: IMedia_Position = Position: End Property
Private Property Get IMedia_Position64() As Int64: IMedia_Position64 = Position64: End Property
Private Sub IMedia_SeekMedia(ByVal Origin As SeekOrigin, Offset As Int64): End Sub
Private Property Get IMedia_Attributes() As MediaGenericAttributes: Set IMedia_Attributes = Attributes: End Property
'
'</section>
'---------------------------------------------
'---------------------------------------------
