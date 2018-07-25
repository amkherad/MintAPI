VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MIDIInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "MIDIInput provided to system MIDI input device.\r\n\r\nIDevice<MIDIInput>\r\nIMediaDevice<MIDIInput>\r\nIInputDevice<MIDIInput>"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Base 0
Option Explicit
Const CLASSID As String = "MIDIInput"
Const CLASSALIAS As String = APPLICATIONDOMAIN & "." & CLASSID
Const CLASSVERSIONSTRING As String = "0.1"
Const CLASSALIASSTRING As String = CLASSALIAS & "/" & CLASSVERSIONSTRING


Implements IDevice
Implements IInputDevice
Implements IMediaDevice
Implements IAliasable

Dim hndl As Long

Friend Sub Initialize()

End Sub

Public Function Listen(Optional NumberOfNotesToRead, Optional ByVal Timeout As Long = tINFINITE) As NoteBuffer

End Function
Public Function ConcurrentListen(Optional NumberOfNotesToRead, Optional ByVal Timeout As Long = tINFINITE) As Thread

End Function
Public Sub StopListening()

End Sub

Private Property Get IDevice_DeviceType() As DeviceType: IDevice_DeviceType = dtDevice_MIDIInput: End Property
Private Property Get IDevice_Handle() As Long: IDevice_Handle = hndl: End Property

Private Sub IInputDevice_ReadInput(Data As Variant)

End Sub

Private Property Get IMediaDevice_DeviceType() As DeviceType: IMediaDevice_DeviceType = dtDevice_MIDIInput: End Property
Private Property Get IMediaDevice_Handle() As Long: IMediaDevice_Handle = hndl: End Property

Private Property Get IInputDevice_DeviceType() As DeviceType: IInputDevice_DeviceType = dtDevice_MIDIInput: End Property
Private Property Get IInputDevice_Handle() As Long: IInputDevice_Handle = hndl: End Property

Public Property Get DeviceType() As DeviceType: DeviceType = dtDevice_MIDIInput: End Property
Public Property Get Handle() As Long: Handle = hndl: End Property


Private Property Get IAliasable_Alias() As String
    IAliasable_Alias = CLASSALIASSTRING
End Property
