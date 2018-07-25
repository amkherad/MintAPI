VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IMedia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'@PROJECT_LICENSE

''<class name="IMedia" interface tags="">
''  <summary>
''    Some class description.
''  </summary>
''  <using>Some class alias.</using>
''  <idea>Me</idea>
''  <sample></sample>
''</class>

Option Base 0
Option Explicit

'=============================================
'=============================================
'=============================================
'<section Meta Data>
'
Const CLASSID As String = "IMedia" '#DO NOT REMOVE IT
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
''<remarks>
''  Bits on {0000 1011 1111 1111 1111 1110 0100 0000} are free to use by users (programmers).
''</remarks>
Public Enum MediaStreamState '@Flag
    mssNone = 0
    mssOpen = 1
    mssPlaying = 3
    mssPaused = &H4 Or mssOpen
    'mssSuspended = mssPaused
    mssStopped = &H8 Or mssPaused Or mssOpen
    mssSeek = &H10
    mssClosed = &HC0000000
    mssCustom = &H100
    mssError = &H20000000
End Enum

'comments comes from MSDN::Multimedia::set
Public Enum MediaTimingFormat '@Sequence
    mtfBytes
    mtfMilliseconds
    mtfFrames
    ''<summary>
    ''  Sets the time format to milliseconds. All commands that use position values will assume milliseconds..
    ''  <br>
    ''  For sequencer devices, the sequence file sets the default format to PPQN or SMPTE. Video-overlay devices do not support this flag.
    ''</summary>
    mtfHMS
    ''<summary>
    ''  Sets the time format to minutes, seconds, and frames. All commands that use position values will assume MSF (the default format for CD audio).<br>
    ''  Specify an MSF value as mm:ss:ff, where mm is minutes, ss is seconds, and ff is frames. You can omit a field if it and all following fields are zero. For example, 3, 3:0, and 3:0:0 are valid ways to express 3 minutes.
    ''  <br><br>
    ''  The MSF fields have the following maximum values:<br>
    ''   Minutes 99<br>
    ''   Seconds 59<br>
    ''   Frames 74
    ''</summary>
    mtfMSF 'minutes, seconds, and frames
    mtfTrack
    mtfSamples
    ''<summary>
    ''  Sets the time format to song pointer (sixteenth notes). All commands that use position values will assume song pointer units. This flag is valid only for a sequence of division type PPQN.
    ''</summary>
    mtfSongPointer
    ''<summary>
    ''  Sets the time format to tracks, minutes, seconds, and frames. All commands that use position values will assume TMSF.
    ''<br>
    ''  Specify a TMSF value as tt:mm:ss:ff, where tt is tracks, mm is minutes, ss is seconds, and ff is frames. You can omit a field if it and all following fields are zero. For example, 3, 3:0, 3:0:0, and 3:0:0:0 are all valid ways to express track 3.
    ''<br><br>
    ''  The TMSF fields have the following maximum values:
    ''<br><br>
    ''  Tracks 99<br>
    ''  Minutes 99<br>
    ''  Seconds 59<br>
    ''  Frames 74
    ''</summary>
    mtfTMSF
    mtfCounter 'Sets the position-information mode to use the VCR counters.
    mtfTimeCode 'Sets the position information mode to use "timecode" information on the tape.
    mtfDetect
    mtfSMPTE
    mtfSMPTE24
    mtfSMPTE25
    mtfSMPTE30
    ''<summary>
    ''  Sets the time format to SMPTE 30 drop frame rate.
    ''<br>
    ''  For VCRs, same as SMPTE 30, except that certain timecode positions are dropped from the format to have the recorded timecode positions for each frame (at the NTSC frame rate of 29.97 fps) correspond to real time (at 30 fps). Timecode positions that are dropped are as follows: two every minute, on the minute, for the first nine of every ten minutes of recorded content. For example, at 01:04:59:29, the next timecode position would be 01:05:00:02, not 01:05:00:00.
    ''<br><br>
    ''  All commands that use position values will assume SMPTE format.
    ''<br><br>
    ''  The sequence file sets the default format to PPQN or SMPTE.
    ''</summary>
    mtfSMPTE30Drop
    mtfDefault
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

Public Property Get Handle() As Long: End Property

Public Property Get Stream() As Stream: End Property

Public Property Get IsOnStorage() As Boolean: End Property

Public Property Get Path() As String: End Property

Public Property Get State() As MediaStreamState: End Property

Public Property Get Length() As Long: End Property
Public Property Get Length64() As Int64: End Property
Public Property Get Position() As Long: End Property
Public Property Get Position64() As Int64: End Property

Public Sub SeekMedia(ByVal Origin As SeekOrigin, Offset As Int64): End Sub

Public Property Get Attributes() As MediaGenericAttributes: End Property

'*********************************************
'*********************************************
'
'</section>
'---------------------------------------------
'---------------------------------------------

