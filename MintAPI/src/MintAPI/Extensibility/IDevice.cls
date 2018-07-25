VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IDevice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"


'Ideas from Me , RoNOTE MIDI Engine (my work!)

Option Base 0
Option Explicit
Const CLASSID As String = "IDevice"
Const CLASSALIAS As String = APPLICATIONDOMAIN & "." & CLASSID
Const CLASSVERSIONSTRING As String = "0.1"
Const CLASSALIASSTRING As String = CLASSALIAS & "/" & CLASSVERSIONSTRING

'================================::Reserved Device Types::================================
'all must multiply by dtDeviceShift
'Device ID     Device Name
'===========
'Group Media
'-----------
'30            MCI
'

Public Enum DeviceType
    '[2  bit:Direction]
    '[8  bit:Group]
    '[15 bit:DeviceID]
    '[7  bit:Free (Used for Flags)]
'=====================================================================
                                                   '3    0    0    0    0    0    0    0
dtDirectionMask = &HC0000000                       '1100 0000 0000 0000 0000 0000 0000 0000
    
    dtOutputDevice = &H40000000                    '0100 0000 0000 0000 0000 0000 0000 0000
    dtInputDevice = &H80000000                     '1000 0000 0000 0000 0000 0000 0000 0000
'=====================================================================
                                                   '   3  F    C
dtGroupMask = &H3FC00000                           '00[11 1111 11]00 0000 0000 0000 0000 0000
dtGroupShift = &H400000

    dtGroup_IO = &H400000                          '00[00 0000 01]00
    dtGroup_Media = &H800000                       '00[00 0000 10]00
    dtGroup_Audio = dtGroup_Media                  'Copy
    dtGroup_Video = dtGroup_Media                  'Copy
    dtGroup_Provider = &HC00000                    '00[00 0000 11]00
    dtGroup_Hardware = &H1000000                   '00[00 0001 00]00
    dtGroup_Graphic = &H1400000                    '00[00 0001 01]00
    dtGroup_Storage = &H1800000                    '00[00 0001 10]00
    
    dtGroup_Others = &H3FC00000                    '00[11 1111 11]00
    dtGroup_Others1 = &H3F800000                   '00[11 1111 10]00
'=====================================================================
                                                   '003FFF80
                                                   '0  3  F    F    F    8 0   0
dtDeviceMask = &H3FFF800                           '00[11 1111 1111 1111 1]000 0000
dtDeviceShift = &H80                               '00[00 0000 0000 0000 1]000 0000
'---------------------------------------------------------------------
    dtDevice_MIDI = dtGroup_Media Or (&H10000)     '00 00 0001 0000 0000 0 000  0000
    dtDevice_MIDIOutput = dtOutputDevice Or dtDevice_MIDI
    dtDevice_MIDIInput = dtInputDevice Or dtDevice_MIDI
'---------------------------------------------------------------------
                                                   '00 00 0001 0000 0000 0 000  0000
    dtDevice_File = dtGroup_IO Or (&H10000) Or dtOutputDevice Or dtInputDevice
'---------------------------------------------------------------------
                                                   '00 00 0010 0000 0000 0 000  0000
    dtDevice_Socket = dtGroup_IO Or (&H20000) Or dtOutputDevice Or dtInputDevice
'---------------------------------------------------------------------
                                                   '00 00 0011 0000 0000 0 000  0000
    dtDevice_Console = dtGroup_IO Or (&H30000) Or dtOutputDevice Or dtInputDevice
'---------------------------------------------------------------------
                                                   '00 00 0100 0000 0000 0 000  0000
    dtDevice_Process = dtGroup_IO Or (&H40000) Or dtOutputDevice Or dtInputDevice
'---------------------------------------------------------------------
'=====================================================================
    
    dtFlags_Obsolete = &H1
    dtFlags_External = &H2
    dtFlags_Plugin = &H4
    dtFlags_Providable = &H8
End Enum


Public Property Get Handle() As Long: End Property
Public Property Get DeviceType() As DeviceType: End Property
