VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MIDIOutput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "MIDIOutput provided to system MIDI output device.\r\n\r\nIDevice<MIDIOutput>\r\nIMediaDevice<MIDIOutput>\r\nIOutputDevice<MIDIOutput>"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Base 0
Option Explicit
Const CLASSID As String = "MIDIOutput"
Const CLASSALIAS As String = APPLICATIONDOMAIN & "." & CLASSID
Const CLASSVERSIONSTRING As String = "0.1"
Const CLASSALIASSTRING As String = CLASSALIAS & "/" & CLASSVERSIONSTRING

'C4 : 440hz

Const SIGNAL_PLAY As String = "play"
Const SIGNAL_STOP As String = "stop"
Const SIGNAL_STATUSCHANGED As String = "statuschanged"
Const SIGNAL_PROGRESS As String = "progress"

Const SLOT_MUTE As String = "mute"
Const SLOT_STOP As String = "stop"
Const SLOT_PLAY As String = "play"


Implements IDevice
Implements IOutputDevice
Implements IMediaDevice
Implements IAliasable

Public Type MIDIMessage
    Volume As Long
    Frequence As Long
    Pitch As Long
End Type

Dim hndl As Long

Friend Sub Initialize()
    
End Sub

Public Sub trigChannel()
    
End Sub
Public Function EnumDevices() As String
    
End Function
Public Function CountDevices() As Long
    
End Function


Public Sub SetVolume(Value As Long)
    
End Sub

Public Sub BeginBend()
    
End Sub
Public Sub EndBend()
    
End Sub
Public Sub BeginSlide()
    
End Sub
Public Sub EndSlide()
    
End Sub
Public Sub BeginVibration()
    
End Sub
Public Sub EndVibration()
    
End Sub
Public Sub SetChorus()
    
End Sub
Public Sub SetMiddle()
    
End Sub
Public Sub SetReverb()
    
End Sub
Public Sub SetDelay()
    
End Sub
Public Sub SetFlanger()
    
End Sub
Public Sub SetPhaser()
    
End Sub
Public Sub EffectWah()
    
End Sub


Public Sub SetMute(Optional Mute As Boolean = True)
    
End Sub


Public Sub SetMessageState(MM As MIDIMessage, Optional State As Boolean = True)
    
End Sub

Public Sub PlayMIDI(nBuffer As NoteBuffer)
    
End Sub
Public Function ConcurrentPlayMIDI(nBuffer As NoteBuffer) As Thread
    
End Function
Public Sub StopMIDI()
    
End Sub

Private Property Get IDevice_DeviceType() As DeviceType: IDevice_DeviceType = dtDevice_MIDIOutput: End Property
Private Property Get IDevice_Handle() As Long: IDevice_Handle = hndl: End Property

Private Property Get IMediaDevice_DeviceType() As DeviceType: IMediaDevice_DeviceType = dtDevice_MIDIOutput: End Property
Private Property Get IMediaDevice_Handle() As Long: IMediaDevice_Handle = hndl: End Property

Private Property Get IOutputDevice_DeviceType() As DeviceType: IOutputDevice_DeviceType = dtDevice_MIDIOutput: End Property
Private Property Get IOutputDevice_Handle() As Long: IOutputDevice_Handle = hndl: End Property

Public Property Get DeviceType() As DeviceType: DeviceType = dtDevice_MIDIOutput: End Property
Public Property Get Handle() As Long: Handle = hndl: End Property

Private Sub IOutputDevice_WriteOutput(Data As Variant)
    
End Sub

Private Property Get IAliasable_Alias() As String
    IAliasable_Alias = CLASSALIASSTRING
End Property
