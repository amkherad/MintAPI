VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MCIDevices"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'@PROJECT_LICENSE

''<class name="MCIDevices" static tags="">
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
Const CLASSID As String = "MCIDevices" '#DO NOT REMOVE IT
Const CLASSALIAS As String = APPLICATIONDOMAIN & "." & CLASSID '#DO NOT REMOVE IT
Const CLASSVERSIONSTRING As String = "0.1" '#DO NOT REMOVE IT
Const CLASSALIASSTRING As String = CLASSALIAS & "/" & CLASSVERSIONSTRING '#DO NOT REMOVE IT

Const DEVNAME_CDAUDIO               As String = "CDAudio"
Const DEVNAME_AVIVIDEO              As String = "AVIVideo"
Const DEVNAME_MPEGVIDEO             As String = "MPEGVideo"
Const DEVNAME_SEQUENCER             As String = "Sequencer"
Const DEVNAME_WAVEAUDIO             As String = "WaveAudio"
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
'<section Creators>
'
Public Function Create(ByVal Name As String, Optional ByVal Flags As Long) As MCIDevice
    Set Create = New MCIDevice
    Call Create.Constructor0(Name, Flags)
End Function

Public Function Auto() As MCIDevice
    Set Auto = Create("", 0)
End Function

Public Function CDAudio() As MCIDevice
    Set CDAudio = Create(DEVNAME_CDAUDIO, 0)
End Function
Public Function AVIVideo() As MCIDevice
    Set AVIVideo = Create(DEVNAME_AVIVIDEO, 0)
End Function
Public Function MPEGVideo() As MCIDevice
    Set MPEGVideo = Create(DEVNAME_MPEGVIDEO, 0)
End Function
Public Function Sequencer() As MCIDevice
    Set Sequencer = Create(DEVNAME_SEQUENCER, 0)
End Function
Public Function WaveAudio() As MCIDevice
    Set WaveAudio = Create(DEVNAME_WAVEAUDIO, 0)
End Function
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

'All your methods here...

'*********************************************
'*********************************************
'
'</section>
'---------------------------------------------
'---------------------------------------------
