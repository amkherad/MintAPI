Attribute VB_Name = "modMCI"
Option Explicit

Public Const MCI_REGPATH_GLOBAL As String = "HKEY_LOCAL_MACHINE\Software\Microsoft\Windows NT\CurrentVersion\"

Private Declare Function API_mciExecute Lib "winmm" Alias "mciExecute" (ByVal lpstrCommand As String) As Long
Private Declare Function API_mciGetCreatorTask Lib "winmm" Alias "mciGetCreatorTask" (ByVal wDeviceID As Long) As Long
Private Declare Function API_mciGetDeviceID Lib "winmm" Alias "mciGetDeviceIDA" (ByVal lpstrName As String) As Long
Private Declare Function API_mciGetDeviceIDFromElementID Lib "winmm" Alias "mciGetDeviceIDFromElementIDA" (ByVal dwElementID As Long, ByVal lpstrType As String) As Long
Private Declare Function API_mciGetErrorString Lib "winmm" Alias "mciGetErrorStringA" (ByVal dwError As Long, lpstrBuffer As Any, ByVal uLength As Long) As Boolean
Private Declare Function API_mciGetYieldProc Lib "winmm" Alias "mciGetYieldProc" (ByVal mciId As Long, pdwYieldData As Long) As Long
Private Declare Function API_mciSendCommand Lib "winmm" Alias "mciSendCommandA" (ByVal wDeviceID As Long, ByVal uMessage As Long, ByVal dwParam1 As Long, ByVal dwParam2 As Long) As Long
Private Declare Function API_mciSendString Lib "winmm" Alias "mciSendStringA" (ByVal lpstrCommand As String, lpstrReturnString As Any, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Private Declare Function API_mciSetYieldProc Lib "winmm" Alias "mciSetYieldProc" (ByVal mciId As Long, ByVal fpYieldProc As Long, ByVal dwYieldData As Long) As Long

'/* flags for dwFlags parameter of MCI_SET command message */
Public Const MCI_SET_DOOR_OPEN                  As Long = &H100
Public Const MCI_SET_DOOR_CLOSED                As Long = &H200
Public Const MCI_SET_TIME_FORMAT                As Long = &H400
Public Const MCI_SET_AUDIO                      As Long = &H800
Public Const MCI_SET_VIDEO                      As Long = &H1000
Public Const MCI_SET_ON                         As Long = &H2000
Public Const MCI_SET_OFF                        As Long = &H4000


Private Const MCI_STRING_OFFSET = 512
Private Const MCI_VD_OFFSET = 1024
Private Const MCI_CD_OFFSET = 1088
Private Const MCI_WAVE_OFFSET = 1152
Private Const MCI_SEQ_OFFSET = 1216

Public Enum MCI_COMMANDS
    MCI_OPEN = &H803
    MCI_CLOSE = &H804
    MCI_ESCAPE = &H805
    MCI_PLAY = &H806
    MCI_SEEK = &H807
    MCI_STOP = &H808
    MCI_PAUSE = &H809
    MCI_INFO = &H80A
    MCI_GETDEVCAPS = &H80B
    MCI_SPIN = &H80C
    MCI_SET = &H80D
    MCI_STEP = &H80E
    MCI_RECORD = &H80F
    MCI_SYSINFO = &H810
    MCI_BREAK = &H811
    MCI_SAVE = &H813
    MCI_STATUS = &H814
    MCI_CUE = &H830
    MCI_REALIZE = &H840
    MCI_WINDOW = &H841
    MCI_PUT = &H842
    MCI_WHERE = &H843
    MCI_FREEZE = &H844
    MCI_UNFREEZE = &H845
    MCI_LOAD = &H850
    MCI_CUT = &H851
    MCI_COPY = &H852
    MCI_PASTE = &H853
    MCI_UPDATE = &H854
    MCI_RESUME = &H855
    MCI_DELETE = &H856
End Enum

Public Enum MCI_MODES
    '/* return values for 'status mode' command */
    MCI_MODE_NOT_READY = (MCI_STRING_OFFSET + 12)
    MCI_MODE_STOP = (MCI_STRING_OFFSET + 13)
    MCI_MODE_PLAY = (MCI_STRING_OFFSET + 14)
    MCI_MODE_RECORD = (MCI_STRING_OFFSET + 15)
    MCI_MODE_SEEK = (MCI_STRING_OFFSET + 16)
    MCI_MODE_PAUSE = (MCI_STRING_OFFSET + 17)
    MCI_MODE_OPEN = (MCI_STRING_OFFSET + 18)
End Enum

Public Const MCI_COMMAND_LENGTH As Long = 127
Public Const MCI_CDAUDIO_DEVICENAME As String = "cdaudio"

Public Const MCI_FUNC_OPEN      As String = "open"
Public Const MCI_FUNC_STOP      As String = "stop"
Public Const MCI_FUNC_PLAY      As String = "play"
Public Const MCI_FUNC_PAUSE     As String = "pause"
Public Const MCI_FUNC_RESUME    As String = "resume"
Public Const MCI_FUNC_STATUS    As String = "status"
Public Const MCI_FUNC_SEEK      As String = "seek"
Public Const MCI_FUNC_CLOSE     As String = "close"

Public Const MCI_FUNC_SETAUDIO  As String = "setaudio"

Public Enum MCI_DEVICE_TYPES
    '/* device ID for "all devices" */
    MCI_ALL_DEVICE_ID = -1 'Any device
    
    '/* constants for predefined MCI device types */
    MCI_DEVTYPE_VCR = 513                'Video-cassette recorder   /* (MCI_STRING_OFFSET + 1) */
    MCI_DEVTYPE_VIDEODISC = 514          'Videodisc player          /* (MCI_STRING_OFFSET + 2) */
    MCI_DEVTYPE_OVERLAY = 515            'Video-overlay device      /* (MCI_STRING_OFFSET + 3) */
    MCI_DEVTYPE_CD_AUDIO = 516           'CD audio device           /* (MCI_STRING_OFFSET + 4) */
    MCI_DEVTYPE_DAT = 517                'Digital-audio tape device /* (MCI_STRING_OFFSET + 5) */
    MCI_DEVTYPE_SCANNER = 518            'Scanner device            /* (MCI_STRING_OFFSET + 6) */
    MCI_DEVTYPE_ANIMATION = 519          'Animation-playback device /* (MCI_STRING_OFFSET + 7) */
    MCI_DEVTYPE_DIGITAL_VIDEO = 520      'Digital-video playback device /* (MCI_STRING_OFFSET + 8) */
    MCI_DEVTYPE_OTHER = 521              'Undefined device          /* (MCI_STRING_OFFSET + 9) */
    MCI_DEVTYPE_WAVEFORM_AUDIO = 522     'Waveform-audio device     /* (MCI_STRING_OFFSET + 10) */
    MCI_DEVTYPE_SEQUENCER = 523          'MIDI sequencer device     /* (MCI_STRING_OFFSET + 11) */
    
    MCI_DEVTYPE_FIRST = MCI_DEVTYPE_VCR
    MCI_DEVTYPE_LAST = MCI_DEVTYPE_SEQUENCER
End Enum

Public Enum MCI_EXTRA_FLAGS
'/* common flags for dwFlags parameter of MCI command messages */
    MCI_NOTIFY = 1
    MCI_WAIT = 2
    MCI_FROM = 4
    MCI_TO = 8
    MCI_TRACK = 10
End Enum

Public Enum MCI_OPEN_FLAGS
    MCI_OPEN_ALIAS = &H400 'An alias is included in the lpstrAlias member of the structure identified by lpOpen.
    MCI_OPEN_SHAREABLE = &H100 'The device or file should be opened as sharable.
    MCI_OPEN_TYPE = &H2000 'A device type name or constant is included in the lpstrDeviceType member of the structure identified by lpOpen.
    MCI_OPEN_TYPE_ID = &H1000 'The low-order word of the lpstrDeviceType member of the structure identified by lpOpen contains a standard MCI device type identifier and the high-order word optionally contains the ordinal index for the device. Use this flag with the MCI_OPEN_TYPE flag.
'The following additional flags apply to compound devices:
    MCI_OPEN_ELEMENT = &H200 'A filename is included in the lpstrElementName member of the structure identified by lpOpen.
    MCI_OPEN_ELEMENT_ID = &H800 'The lpstrElementName member of the structure identified by lpOpen is interpreted as a DWORD value and has meaning internal to the device. Use this flag with the MCI_OPEN_ELEMENT flag.
    '
    'The following additional flags are used with the digitalvideo device type:
    '
    'MCI_DGV_OPEN_NOSTATIC
    '
    '    The device should reduce the number of static (system) colors in the palette. This increases the number of colors available for rendering the video stream. This flag applies only to devices that share a palette with Windows.
    'MCI_DGV_OPEN_PARENT
    '
    '    The parent window handle is specified in the hWndParent member of the structure identified by lpOpen.
    'MCI_DGV_OPEN_WS
    '
    '    A window style is specified in the dwStyle member of the structure identified by lpOpen.
    'MCI_DGV_OPEN_16BIT
    '
    '    Indicates a preference for 16-bit MCI device support.
    'MCI_DGV_OPEN_32BIT
    '
    '    Indicates a preference for 32-bit MCI device support.
    '
    'For digital-video devices, the lpOpen parameter points to an MCI_DGV_OPEN_PARMS structure.
    '
    'The following additional flags are used with the overlay device type:
    '
    'MCI_OVLY_OPEN_PARENT
    '
    '    The parent window handle is specified in the hWndParent member of the structure identified by lpOpen.
    'MCI_OVLY_OPEN_WS
    '
    '    A window style is specified in the dwStyle member of the structure identified by lpOpen. The dwStyle value specifies the style of the window that the driver will create and display if the application does not provide one. The style parameter takes an integer that defines the window style. These constants are the same as the standard window styles (such as WS_CHILD, WS_OVERLAPPEDWINDOW, or WS_POPUP).
    '
    'For video-overlay devices, the lpOpen parameter points to an MCI_OVLY_OPEN_PARMS structure.
    '
    'The following additional flag is used with the waveaudio device type:
    '
    'MCI_WAVE_OPEN_BUFFER
    '
    '    A buffer length is specified in the dwBufferSeconds member of the structure identified by lpOpen.
    '
    'For waveform-audio devices, the lpOpen parameter points to an MCI_WAVE_OPEN_PARMS structure. The MCIWAVE driver requires an asynchronous waveform-audio device.
End Enum
Public Type MCI_OPEN_PARMS
    dwCallback As Long 'The low-order word specifies a window handle used for the MCI_NOTIFY flag.
    wDeviceID As Long 'Identifier returned to application.
    lpstrDeviceType As String 'Name or constant identifier of the device type. (The name of the device is typically obtained from the registry or SYSTEM.INI file.) If this member is a constant, it can be one of the values listed in MCI Device Types.
    lpstrElementName As String 'Device element (often a path).
    lpstrAlias As String 'Optional device alias.
End Type

Public Enum MCI_PLAY_Flags
    MCI_FROM
'
'    A starting location is included in the dwFrom member of the structure identified by lpPlay. The units assigned to the position values are specified with the MCI_SET_TIME_FORMAT flag of the MCI_SET command. If MCI_FROM is not specified, the starting location defaults to the current position.
'MCI_TO
'
'    An ending location is included in the dwTo member of the structure identified by lpPlay. The units assigned to the position values are specified with the MCI_SET_TIME_FORMAT flag of MCI_SET. If MCI_TO is not specified, the ending location defaults to the end of the media.
'
'The following additional flags are used with the digitalvideo device type:
'
'MCI_DGV_PLAY_REPEAT
'
'    Playback should start again at the beginning when the end of the content is reached.
'MCI_DGV_PLAY_REVERSE
'
'    Playback should occur in reverse.
'MCI_MCIAVI_PLAY_WINDOW
'
'    Playback should occur in the window associated with a device instance (the default). (This flag is specific to MCIAVI.DRV.)
'MCI_MCIAVI_PLAY_FULLSCREEN
'
'    Playback should use a full-screen display. Use this flag only when playing compressed or 8-bit files.
'
'For digital-video devices, lpPlay points to an MCI_DGV_PLAY_PARMS structure.
'
'The following additional flags are used with the vcr device type:
'
'MCI_VCR_PLAY_AT
'
'    The dwAt member of the structure identified by lpPlay contains a time when the entire command begins, or if the device is cued, when the device reaches the from position given by the MCI_CUE command.
'MCI_VCR_PLAY_REVERSE
'
'    Playback should occur in reverse.
'MCI_VCR_PLAY_SCAN
'
'    Playback should be as fast as possible while maintaining video output.
'
'For VCR devices, lpPlay points to an MCI_VCR_PLAY_PARMS structure.
'
'The following additional flags are used with the videodisc device type:
'
'MCI_VD_PLAY_FAST
'
'    Play fast.
'MCI_VD_PLAY_REVERSE
'
'    Play in reverse.
'MCI_VD_PLAY_SCAN
'
'    Scan quickly.
'MCI_VD_PLAY_SLOW
'
'    Play slowly.
'MCI_VD_PLAY_SPEED
'
'    The play speed is included in the dwSpeed member in the structure identified by lpPlay.

End Enum
Public Type MCI_PLAY_PARMS
  dwCallback As Long 'The low-order word specifies a window handle used for the MCI_NOTIFY flag.
  dwFrom As Long 'Position to play from.'A starting location is included in the dwFrom member of the structure identified by lpPlay. The units assigned to the position values are specified with the MCI_SET_TIME_FORMAT flag of the MCI_SET command. If MCI_FROM is not specified, the starting location defaults to the current position.
  dwTo As Long 'Position to play to.'An ending location is included in the dwTo member of the structure identified by lpPlay. The units assigned to the position values are specified with the MCI_SET_TIME_FORMAT flag of MCI_SET. If MCI_TO is not specified, the ending location defaults to the end of the media.
End Type

Public Enum MCI_STATUS_FLAGS
    MCI_STATUS_ITEM
        ''Specifies that the dwItem member of the structure identified by lpStatus contains a constant specifying which status item to obtain. The following constants define which status item to return in the dwReturn member of the structure:
    
    MCI_STATUS_CURRENT_TRACK = &H8
    
        ''The dwReturn member is set to the current track number. MCI uses continuous track numbers.
    
    MCI_STATUS_LENGTH = &H1
    
        ''The dwReturn member is set to the total media length.
    MCI_STATUS_MODE = &H4
        ''The dwReturn member is set to the current mode of the device. The modes include the following:
    
        ''    ''MCI_MODE_NOT_READY
        ''    ''MCI_MODE_PAUSE
        ''    ''MCI_MODE_PLAY
        ''    ''MCI_MODE_STOP
        ''    ''MCI_MODE_OPEN
        ''    ''MCI_MODE_RECORD
        ''    ''MCI_MODE_SEEK
    
    MCI_STATUS_NUMBER_OF_TRACKS = &H3
    
        ''The dwReturn member is set to the total number of playable tracks.
    MCI_STATUS_POSITION = &H2
    
        ''The dwReturn member is set to the current position.
    MCI_STATUS_READY = &H7
    
        ''The dwReturn member is set to TRUE if the device is ready; it is set to FALSE otherwise.
    MCI_STATUS_TIME_FORMAT = &H6
    
        ''The dwReturn member is set to the current time format of the device. The time formats include:
    
        ''    ''MCI_FORMAT_BYTES
        ''    ''MCI_FORMAT_FRAMES
        ''    ''MCI_FORMAT_HMS
        ''    ''MCI_FORMAT_MILLISECONDS
        ''    ''MCI_FORMAT_MSF
        ''    ''MCI_FORMAT_SAMPLES
        ''    ''MCI_FORMAT_TMSF
    
    MCI_STATUS_START
    
        ''Obtains the starting position of the media. To get the starting position, combine this flag with MCI_STATUS_ITEM and set the dwItem member of the structure identified by lpStatus to MCI_STATUS_POSITION.
    MCI_TRACK
    
        ''Indicates a status track parameter is included in the dwTrack member of the structure identified by lpStatus. You must use this flag with the MCI_STATUS_POSITION or MCI_STATUS_LENGTH constants. When used with MCI_STATUS_POSITION, MCI_TRACK obtains the starting position of the specified track. When used with MCI_STATUS_LENGTH, MCI_TRACK obtains the length of the specified track. MCI uses continuous track numbers.
    
    'The following additional flags are used with the cdaudio device type. These constants are used in the dwItem member of the structure pointed to by the lpStatus parameter when MCI_STATUS_ITEM is specified for the dwFlags parameter.
    
    MCI_CDA_STATUS_TYPE_TRACK
    
        ''The dwReturn member is set to one of the following values:
    
        ''    ''MCI_CDA_TRACK_AUDIO
        ''    ''MCI_CDA_TRACK_OTHER
    
        ''To use this flag, the MCI_TRACK flag must be set, and the dwTrack member of the structure identified by lpStatus must contain a valid track number.
    MCI_STATUS_MEDIA_PRESENT = &H5
    
        ''The dwReturn member is set to TRUE if the media is inserted in the device; it is set to FALSE otherwise.
    
    'The following additional flags are used with the digitalvideo device type:
    
    'MCI_DGV_STATUS_DISKSPACE
    '
    '    ''The lpstrDrive member of the structure identified by lpStatus specifies a disk drive or, in some implementations, a path. The MCI_STATUS command returns the approximate amount of disk space that could be obtained by the MCI_RESERVE command in the dwReturn member of the structure identified by lpStatus. The disk space is measured in units of the current time format.
    'MCI_DGV_STATUS_INPUT
    '
    '    ''The constant specified by the dwItem member of the structure identified by lpStatus applies to the input.
    'MCI_DGV_STATUS_LEFT
    '
    '    ''The constant specified by the dwItem member of the structure identified by lpStatus applies to the left audio channel.
    'MCI_DGV_STATUS_NOMINAL
    '
    '    ''The constant specified by the dwItem member of the structure identified by lpStatus requests the nominal value rather than the current value.
    'MCI_DGV_STATUS_OUTPUT
    '
    '    ''The constant specified by the dwItem member of the structure identified by lpStatus applies to the output.
    'MCI_DGV_STATUS_RECORD
    '
    '    ''The frame rate returned for the MCI_DGV_STATUS_FRAME_RATE flag is the rate used for compression.
    'MCI_DGV_STATUS_REFERENCE
    '
    '    ''The dwReturn member of the structure identified by lpStatus returns the nearest key-frame image that precedes the frame specified in the dwReference member.
    'MCI_DGV_STATUS_RIGHT
    '
    '    ''The constant specified by the dwItem member of the structure identified by lpStatus applies to the right audio channel.
    '
    ''The following constants are used with the digitalvideo device type in the dwItem member of the structure pointed to by the lpStatus parameter when MCI_STATUS_ITEM is specified for the dwFlags parameter.
    '
    'MCI_AVI_STATUS_AUDIO_BREAKS
    '
    '    ''The dwReturn member returns the number of times the audio portion of the last AVI sequence broke up. The system counts an audio break whenever it attempts to write audio data to the device driver and discovers that the driver has already played all of the available data. This flag is recognized only by the MCIAVI digital-video driver.
    'MCI_AVI_STATUS_FRAMES_SKIPPED
    '
    '    ''The dwReturn member returns the number of frames that were not drawn when the last AVI sequence was played. This flag is recognized only by the MCIAVI digital-video driver.
    'MCI_AVI_STATUS_LAST_PLAY_SPEED
    '
    '    ''The dwReturn member returns a value representing how closely the actual playing time of the last AVI sequence matched the target playing time. The value 1000 indicates that the target time and the actual time were the same. A value of 2000, for example, would indicate that the AVI sequence took twice as long to play as it should have. This flag is recognized only by the MCIAVI digital-video driver.
    'MCI_DGV_STATUS_AUDIO
    '
    '    ''The dwReturn member returns MCI_ON or MCI_OFF depending on the most recent MCI_SET_AUDIO option for the MCI_SET command. It returns MCI_ON if either or both speakers are enabled, and MCI_OFF otherwise.
    'MCI_DGV_STATUS_AUDIO_INPUT
    '
    '    ''The dwReturn member returns the approximate instantaneous audio level of the analog audio signal. A value greater than 1000 implies there is clipping distortion. Some devices can determine this value only while recording audio. This status value has no associated MCI_SET or MCI_SETAUDIO command. This value is related to, but normalized differently from, the waveform-audio command MCI_WAVE_STATUS_LEVEL.
    'MCI_DGV_STATUS_AUDIO_RECORD
    '
    '    ''The dwReturn member returns MCI_ON or MCI_OFF reflecting the state set by the MCI_DGV_SETAUDIO_RECORD flag of the MCI_SETAUDIO command.
    'MCI_DGV_STATUS_AUDIO_SOURCE
    '
    '    ''The dwReturn member returns the current audio digitizer source:
    'MCI_DGV_SETAUDIO_AVERAGE
    '
    '    ''Specifies the average of the left and right audio channels.
    'MCI_DGV_SETAUDIO_LEFT
    '
    '    ''Specifies the left audio channel.
    'MCI_DGV_SETAUDIO_RIGHT
    '
    '    ''Specifies the right audio channel.
    'MCI_DGV_SETAUDIO_STEREO
    '
    '    ''Specifies stereo.
    'MCI_DGV_STATUS_AUDIO_STREAM
    '
    '    ''The dwReturn member returns the current audio-stream number.
    'MCI_DGV_STATUS_AVGBYTESPERSEC
    '
    '    ''The dwReturn member returns the average number of bytes per second used for recording.
    'MCI_DGV_STATUS_BASS
    '
    '    ''The dwReturn member returns the current audio bass level. Use MCI_DGV_STATUS_NOMINAL with this flag to obtain the nominal level.
    'MCI_DGV_STATUS_BITSPERPEL
    '
    '    ''The dwReturn member returns the number of bits per pixel used for saving captured or recorded data.
    'MCI_DGV_STATUS_BITSPERSAMPLE
    '
    '    ''The dwReturn member returns the number of bits per sample the device uses for recording. This applies only to devices supporting the PCM format.
    'MCI_DGV_STATUS_BLOCKALIGN
    '
    '    ''The dwReturn member returns the alignment of data blocks relative to the start of the input waveform.
    'MCI_DGV_STATUS_BRIGHTNESS
    '
    '    ''The dwReturn member returns the current video brightness level. Use MCI_DGV_STATUS_NOMINAL with this flag to obtain the nominal level.
    'MCI_DGV_STATUS_COLOR
    '
    '    ''The dwReturn member returns the current color level. Use MCI_DGV_STATUS_NOMINAL with this flag to obtain the nominal level.
    'MCI_DGV_STATUS_CONTRAST
    '
    '    ''The dwReturn member returns the current contrast level. Use MCI_DGV_STATUS_NOMINAL with this flag to obtain the nominal level.
    'MCI_DGV_STATUS_FILEFORMAT
    '
    '    ''The dwReturn member returns the current file format for recording or saving.
    'MCI_DGV_STATUS_FILE_MODE
    '
    '    ''The dwReturn member returns the state of the file operation:
    '
    '    ''MCI_DGV_FILE_MODE_EDITING
    '
    '    ''Returned during cut, copy, delete, paste, and undo operations.
    '
    '    ''MCI_DGV_FILE_MODE_IDLE
    '
    '    ''Returned when the file is ready for the next operation.
    '
    '    ''MCI_DGV_FILE_MODE_LOADING
    '
    '    ''Returned while the file is being loaded.
    '
    '    ''MCI_DGV_FILE_MODE_SAVING
    '
    '    ''Returned while the file is being saved.
    'MCI_DGV_STATUS_FILE_COMPLETION
    '
    '    ''The dwReturn member returns the estimated percentage a load, save, capture, cut, copy, delete, paste, or undo operation has progressed. (Applications can use this to provide a visual indicator of progress.) This flag is not supported by all digital-video devices.
    'MCI_DGV_STATUS_FORWARD
    '
    '    ''The dwReturn member returns TRUE if the device direction is forward or the device is not playing.
    'MCI_DGV_STATUS_FRAME_RATE
    '
    '    ''The dwReturn member must be used with MCI_DGV_STATUS_NOMINAL, MCI_DGV_STATUS_RECORD, or both. When used with MCI_DGV_STATUS_RECORD, the current frame rate used for recording is returned. When used with both MCI_DGV_STATUS_RECORD and MCI_DGV_STATUS_NOMINAL, the nominal frame rate associated with the input video signal is returned. When used with MCI_DGV_STATUS_NOMINAL, the nominal frame rate associated with the file is returned. In all cases the units are in frames per second multiplied by 1000.
    'MCI_DGV_STATUS_GAMMA
    '
    '    ''The dwReturn member returns the current gamma value. Use MCI_DGV_STATUS_NOMINAL with this flag to obtain the nominal level.
    'MCI_DGV_STATUS_HPAL
    '
    '    ''The dwReturn member returns the ASCII decimal value for the current palette handle. The handle is contained in the low-order word of the returned value.
    'MCI_DGV_STATUS_HWND
    '
    '    ''The dwReturn member returns the ASCII decimal value for the current explicit or default window handle associated with this device driver instance. The handle is contained in the low-order word of the returned value.
    'MCI_DGV_STATUS_KEY_COLOR
    '
    '    ''The dwReturn member returns the current key-color value.
    'MCI_DGV_STATUS_KEY_INDEX
    '
    '    ''The dwReturn member returns the current key-index value.
    'MCI_DGV_STATUS_MONITOR
    '
    '    ''The dwReturn member returns a constant indicating the source of the current presentation. The following constants are defined:
    '
    '    ''MCI_DGV_MONITOR_FILE
    '
    '    ''A file is the source.
    '
    '    ''MCI_DGV_MONITOR_INPUT
    '
    '    ''The input is the source.
    'MCI_DGV_STATUS_MONITOR_METHOD
    '
    '    ''The dwReturn member returns a constant indicating the method used for input monitoring. The following constants are defined:
    '
    '    ''MCI_DGV_METHOD_DIRECT
    '
    '    ''Direct input monitoring.
    '
    '    ''MCI_DGV_METHOD_POST
    '
    '    ''Post-input monitoring.
    '
    '    ''MCI_DGV_METHOD_PRE
    '
    '    ''Pre-input monitoring.
    'MCI_DGV_STATUS_PAUSE_MODE
    '
    '    ''The dwReturn member returns MCI_MODE_PLAY if the device was paused while playing and returns MCI_MODE_RECORD if the device was paused while recording. The command returns MCIERR_NONAPPLICABLE_FUNCTION as an error return if the device is not paused.
    'MCI_DGV_STATUS_SAMPLESPERSECOND
    '
    '    ''The dwReturn member returns the number of samples per second recorded.
    'MCI_DGV_STATUS_SEEK_EXACTLY
    '
    '    ''The dwReturn member returns TRUE or FALSE indicating whether or not the seek exactly format is set. (Applications can set this format by using the MCI_SET command with the MCI_DGV_SET_SEEK_EXACTLY flag.)
    'MCI_DGV_STATUS_SHARPNESS
    '
    '    ''The dwReturn member returns the current sharpness level. Use MCI_DGV_STATUS_NOMINAL with this flag to obtain the nominal level.
    'MCI_DGV_STATUS_SIZE
    '
    '    ''The dwReturn member returns the approximate playback duration of compressed data that the reserved workspace will hold. The duration units are in the current time format. It returns zero if there is no reserved disk space. The size returned is approximate since the precise disk space for compressed data cannot, in general, be predicted until after the data has been compressed.
    'MCI_DGV_STATUS_SMPTE
    '
    '    ''The dwReturn member returns the SMPTE time code associated with the current position in the workspace.
    'MCI_DGV_STATUS_SPEED
    '
    '    ''The dwReturn member returns the current playback speed.
    'MCI_DGV_STATUS_STILL_FILEFORMAT
    '
    '    ''The dwReturn member returns the current file format for the MCI_CAPTURE command.
    'MCI_DGV_STATUS_TINT
    '
    '    ''The dwReturn member returns the current video tint level. Use MCI_DGV_STATUS_NOMINAL with this flag to obtain the nominal level.
    'MCI_DGV_STATUS_TREBLE
    '
    '    ''The dwReturn member returns the current audio treble level. Use MCI_DGV_STATUS_NOMINAL with this flag to obtain the nominal level.
    'MCI_DGV_STATUS_UNSAVED
    '
    '    ''The dwReturn member returns TRUE if there is recorded data in the workspace that might be lost as a result of a MCI_CLOSE, MCI_LOAD, MCI_RECORD, MCI_RESERVE, MCI_CUT, MCI_DELETE, or MCI_PASTE command. The member returns FALSE otherwise.
    'MCI_DGV_STATUS_VIDEO
    '
    '    ''The dwReturn member returns MCI_ON if video is enabled or MCI_OFF if it is disabled.
    'MCI_DGV_STATUS_VIDEO_RECORD
    '
    '    ''The dwReturn member returns MCI_ON or MCI_OFF, reflecting the state set by the MCI_DGV_SETVIDEO_RECORD flag of the MCI_SETVIDEO command.
    'MCI_DGV_STATUS_VIDEO_SOURCE
    '
    '    ''The dwReturn member returns a constant indicating the type of video source set by the MCI_DGV_SETVIDEO_SOURCE flag of the MCI_SETVIDEO command.
    'MCI_DGV_STATUS_VIDEO_SRC_NUM
    '
    '    ''The dwReturn member returns the number within its type of the video-input source currently active.
    'MCI_DGV_STATUS_VIDEO_STREAM
    '
    '    ''The dwReturn member returns the current video-stream number.
    'MCI_DGV_STATUS_VOLUME
    '
    '    ''The dwReturn member returns the average of the volume to the left and right speakers. Use MCI_DGV_STATUS_NOMINAL with this flag to obtain the nominal level.
    'MCI_DGV_STATUS_WINDOW_VISIBLE
    '
    '    ''The dwReturn member returns TRUE if the window is not hidden.
    'MCI_DGV_STATUS_WINDOW_MINIMIZED
    '
    '    ''The dwReturn member returns TRUE if the window is minimized.
    'MCI_DGV_STATUS_WINDOW_MAXIMIZED
    '
    '    ''The dwReturn member returns TRUE if the window is maximized.
    'MCI_STATUS_MEDIA_PRESENT
    '
    '    ''The dwReturn member returns TRUE.
    '
    ''For digital-video devices, the lpStatus parameter points to an MCI_DGV_STATUS_PARMS structure.
    '
    ''The following additional flags are used with the sequencer device type. These constants are used in the dwItem member of the structure pointed to by the lpStatus parameter when MCI_STATUS_ITEM is specified for the dwFlags parameter.
    '
    'MCI_SEQ_STATUS_DIVTYPE
    '
    '    ''The dwReturn member is set to one of the following values indicating the current division type of a sequence:
    '
    '    ''    ''MCI_SEQ_DIV_PPQN
    '    ''    ''MCI_SEQ_DIV_SMPTE_24
    '    ''    ''MCI_SEQ_DIV_SMPTE_25
    '    ''    ''MCI_SEQ_DIV_SMPTE_30
    '    ''    ''MCI_SEQ_DIV_SMPTE_30DROP
    '
    'MCI_SEQ_STATUS_MASTER
    '
    '    ''The dwReturn member is set to the synchronization type used for master operation.
    'MCI_SEQ_STATUS_OFFSET
    '
    '    ''The dwReturn member is set to the current SMPTE offset of a sequence.
    'MCI_SEQ_STATUS_PORT
    '
    '    ''The dwReturn member is set to the MIDI device identifier for the current port used by the sequence.
    'MCI_SEQ_STATUS_SLAVE
    '
    '    ''The dwReturn member is set to the synchronization type used for subordinate operation.
    'MCI_SEQ_STATUS_TEMPO
    '
    '    ''The dwReturn member is set to the current tempo of a MIDI sequence in beats per minute for PPQN files, or frames per second for SMPTE files.
    'MCI_STATUS_MEDIA_PRESENT
    '
    '    ''The dwReturn member is set to TRUE if the media is inserted in the device; it is set to FALSE otherwise.
    '
    ''The following additional flags are used with the vcr device type. These constants are used in the dwItem member of the structure pointed to by the lpStatus parameter when MCI_STATUS_ITEM is specified for the dwFlags parameter.
    '
    'MCI_STATUS_MEDIA_PRESENT
    '
    '    ''The dwReturn member is set to TRUE if the media is inserted in the device; it is set to FALSE otherwise.
    'MCI_VCR_STATUS_ASSEMBLE_RECORD
    '
    '    ''The dwReturn member is set to TRUE if assemble mode is on; it is set to FALSE otherwise.
    'MCI_VCR_STATUS_AUDIO_MONITOR
    '
    '    ''The dwReturn member is set to a constant, indicating the currently selected audio-monitor type.
    'MCI_VCR_STATUS_AUDIO_MONITOR_NUMBER
    '
    '    ''The dwReturn member is set to the number of the currently selected audio-monitor type.
    'MCI_VCR_STATUS_AUDIO_RECORD
    '
    '    ''The dwReturn member is set to TRUE if audio will be recorded when the next record command is given; it is set to FALSE otherwise. If you specify MCI_TRACK in the dwFlags parameter of this command, dwTrack contains the track this inquiry applies to.
    'MCI_VCR_STATUS_AUDIO_SOURCE
    '
    '    ''The dwReturn member is set to a constant, indicating the current audio-source type.
    'MCI_VCR_STATUS_AUDIO_SOURCE_NUMBER
    '
    '    ''The dwReturn member is set to the number of the currently selected audio-source type.
    'MCI_VCR_STATUS_CLOCK
    '
    '    ''The dwReturn member is set to the current clock value, in total clock increments.
    'MCI_VCR_STATUS_CLOCK_ID
    '
    '    ''The dwReturn member is set to a number which uniquely describes the clock in use.
    'MCI_VCR_STATUS_COUNTER_FORMAT
    '
    '    ''The dwReturn member is set to a constant describing the current counter format. For more information, see the MCI_SET_TIME_FORMAT flag of the MCI_SET command.
    'MCI_VCR_STATUS_COUNTER_RESOLUTION
    '
    '    ''The dwReturn member is set to a constant describing the resolution of the counter, and is one of the following values:
    '
    '    ''    ''MCI_VCR_COUNTER_RES_FRAMES: Counter has resolution of frames.
    '    ''    ''MCI_VCR_COUNTER_RES_SECONDS: Counter has resolution of seconds.
    '    ''    ''MCI_VCR_STATUS_COUNTER_VALUE: The dwReturn member is set to the current counter reading, in the current counter-time format.
    '
    'MCI_VCR_STATUS_FRAME_RATE
    '
    '    ''The dwReturn member is set to the current native frame rate of the device.
    'MCI_VCR_STATUS_INDEX
    '
    '    ''The dwReturn member is set to a constant, describing the current contents of the on-screen display, and is one of the following:
    '
    '    ''    ''MCI_VCR_INDEX_COUNTER
    '    ''    ''MCI_VCR_INDEX_DATE
    '    ''    ''MCI_VCR_INDEX_TIME
    '    ''    ''MCI_VCR_INDEX_TIMECODE
    '
    'MCI_VCR_STATUS_INDEX_ON
    '
    '    ''The dwReturn member is set to TRUE if the on-screen display is on; it is set to FALSE otherwise.
    'MCI_VCR_STATUS_MEDIA_TYPE
    '
    '    ''The dwReturn member is set to one of the following:
    '
    '    ''    ''MCI_VCR_MEDIA_8MM
    '    ''    ''MCI_VCR_MEDIA_HI8
    '    ''    ''MCI_VCR_MEDIA_VHS
    '    ''    ''MCI_VCR_MEDIA_SVHS
    '    ''    ''MCI_VCR_MEDIA_BETA
    '    ''    ''MCI_VCR_MEDIA_EDBETA
    '    ''    ''MCI_VCR_MEDIA_OTHER
    '
    'MCI_VCR_STATUS_NUMBER
    '
    '    ''The dwNumber member is set to the logical-tuner number when you use this flag with the MCI_VCR_STATUS_TUNER_CHANNEL flag.
    'MCI_VCR_STATUS_NUMBER_OF_AUDIO_TRACKS
    '
    '    ''The dwReturn member is set to the number of audio tracks that are independently selectable.
    'MCI_VCR_STATUS_NUMBER_OF_VIDEO_TRACKS
    '
    '    ''The dwReturn member is set to the number of video tracks that are independently selectable.
    'MCI_VCR_STATUS_PAUSE_TIMEOUT
    '
    '    ''The dwReturn member is set to the maximum duration, in milliseconds, of a pause command. The return value of zero indicates that no time-out will occur.
    'MCI_VCR_STATUS_PLAY_FORMAT
    '
    '    ''The dwReturn member is set to one of the following:
    '
    '    ''    ''MCI_VCR_FORMAT_EP
    '    ''    ''MCI_VCR_FORMAT_LP
    '    ''    ''MCI_VCR_FORMAT_OTHER
    '    ''    ''MCI_VCR_FORMAT_SP
    '
    'MCI_VCR_STATUS_POSTROLL_DURATION
    '
    '    ''The dwReturn member is set to the length of the videotape that will play after the spot at which it was stopped, in the current time format. This is needed to brake the VCR tape transport from a stop or pause command.
    'MCI_VCR_STATUS_POWER_ON
    '
    '    ''The dwReturn member is set to TRUE if the power is on; it is set to FALSE otherwise.
    'MCI_VCR_STATUS_PREROLL_DURATION
    '
    '    ''The dwReturn member is set to the length of the videotape that will play before the spot at which it was started, in the current time format. This is needed to stabilize the VCR output.
    'MCI_VCR_STATUS_RECORD_FORMAT
    '
    '    ''The dwReturn member is set to one of the following:
    '
    '    ''    ''MCI_VCR_FORMAT_EP
    '    ''    ''MCI_VCR_FORMAT_LP
    '    ''    ''MCI_VCR_FORMAT_OTHER
    '    ''    ''MCI_VCR_FORMAT_SP
    '
    'MCI_VCR_STATUS_SPEED
    '
    '    ''The dwReturn member is set to the current speed. For more information, see the MCI_VCR_SET_SPEED flag of the MCI_SET command.
    'MCI_VCR_STATUS_TIME_MODE
    '
    '    ''The dwReturn member is set to one of the following:
    '
    '    ''    ''MCI_VCR_TIME_COUNTER
    '    ''    ''MCI_VCR_TIME_DETECT
    '    ''    ''MCI_VCR_TIME_TIMECODE
    '
    '    ''For more information, see the MCI_VCR_SET_TIME_MODE flag of the MCI_SET command.
    'MCI_VCR_STATUS_TIME_TYPE
    '
    '    ''The dwReturn member is set to a constant describing the current time type in use (used by play, record, seek, and so on), and is one of the following:
    'MCI_VCR_TIME_COUNTER
    '
    '    ''Counter is in use.
    'MCI_VCR_TIME_TIMECODE
    '
    '    ''Timecode is in use.
    'MCI_VCR_STATUS_TIMECODE_PRESENT
    '
    '    ''The dwReturn member is set to TRUE if timecode is present at the current position in the content; it is set to FALSE otherwise.
    'MCI_VCR_STATUS_TIMECODE_RECORD
    '
    '    ''The dwReturn member is set to TRUE if the timecode will be recorded when the next record command is given; it is set to FALSE otherwise.
    'MCI_VCR_STATUS_TIMECODE_TYPE
    '
    '    ''The dwReturn member is set to a constant, describing the type of timecode that is directly supported by the device, and is one of the following:
    '
    '    ''    ''MCI_VCR_TIMECODE_TYPE_NONE: This device does not use a timecode.
    '    ''    ''MCI_VCR_TIMECODE_TYPE_OTHER: This device uses an unspecified timecode.
    '    ''    ''MCI_VCR_TIMECODE_TYPE_SMPTE: This device uses SMPTE timecode.
    '    ''    ''MCI_VCR_TIMECODE_TYPE_SMPTE_DROP: This device uses SMPTE drop timecode.
    '
    'MCI_VCR_STATUS_TUNER_CHANNEL
    '
    '    ''The dwReturn member is set to the current channel number. If you specify MCI_VCR_STATUS_NUMBER in the dwFlags parameter of this command, dwNumber contains the logical-tuner number this command applies to.
    'MCI_VCR_STATUS_VIDEO_MONITOR
    '
    '    ''The dwReturn member is set to a constant, indicating the currently selected video-monitor type.
    'MCI_VCR_STATUS_VIDEO_MONITOR_NUMBER
    '
    '    ''The dwReturn member is set to the number of the currently selected video-monitor type.
    'MCI_VCR_STATUS_VIDEO_RECORD
    '
    '    ''The dwReturn member is set to TRUE if video will be recorded when the next record command is given; it is set to FALSE otherwise. If you specify MCI_TRACK in the dwFlags parameter of this command, dwTrack contains the track this inquiry applies to.
    'MCI_VCR_STATUS_VIDEO_SOURCE
    '
    '    ''The dwReturn member is set to a constant indicating the currently selected video-source type.
    'MCI_VCR_STATUS_VIDEO_SOURCE_NUMBER
    '
    '    ''The dwReturn member is set to the number of the currently selected video-source type.
    'MCI_VCR_STATUS_WRITE_PROTECTED
    '
    '    ''The dwReturn member is set to TRUE if the media is write-protected; it is set to FALSE otherwise.
    '
    ''For VCR devices, the lpStatus parameter points to an MCI_VCR_STATUS_PARMS structure.
    '
    ''Using the MCI_STATUS_LENGTH flag to determine the length of the media always returns 2 hours for VCR devices, unless the length has been explicitly changed using the MCI_SET command.
    '
    ''The following additional flags are used with the overlay device type. These constants are used in the dwItem member of the structure pointed to by the lpStatus parameter when MCI_STATUS_ITEM is specified for the dwFlags parameter.
    '
    'MCI_OVLY_STATUS_HWND
    '
    '    ''The dwReturn member is set to the handle of the window associated with the video-overlay device.
    'MCI_OVLY_STATUS_STRETCH
    '
    '    ''The dwReturn member is set to TRUE if stretching is enabled; it is set to FALSE otherwise.
    'MCI_STATUS_MEDIA_PRESENT
    '
    '    ''The dwReturn member is set to TRUE if the media is inserted in the device; it is set to FALSE otherwise.
    '
    ''The following additional flags are used with the videodisc device type. These constants are used in the dwItem member of the structure pointed to by the lpStatus parameter when MCI_STATUS_ITEM is specified for the dwFlags parameter.
    '
    'MCI_STATUS_MEDIA_PRESENT
    '
    '    ''The dwReturn member is set to TRUE if the media is inserted in the device; it is set to FALSE otherwise.
    'MCI_STATUS_MODE
    '
    '    ''The dwReturn member is set to the current mode of the device. Videodisc devices can return the MCI_VD_MODE_PARK constant, in addition to the constants any device can return, as documented with the dwFlags parameter.
    'MCI_VD_STATUS_DISC_SIZE
    '
    '    ''The dwReturn member is set to the size of the loaded disc in inches (8 or 12).
    'MCI_VD_STATUS_FORWARD
    '
    '    ''The dwReturn member is set to TRUE if playing forward; it is set to FALSE otherwise.
    '
    '    ''The MCI videodisc device does not support this flag.
    'MCI_VD_STATUS_MEDIA_TYPE
    '
    '    ''The dwReturn member is set to the media type of the inserted media. The following media types can be returned:
    '
    '    ''MCI_VD_MEDIA_CAV
    '
    '    ''MCI_VD_MEDIA_CLV
    '
    '    ''MCI_VD_MEDIA_OTHER
    'MCI_VD_STATUS_SIDE
    '
    '    ''The dwReturn member is set to 1 or 2 to indicate which side of the disc is loaded. Not all videodisc devices support this flag.
    'MCI_VD_STATUS_SPEED
    '
    '    ''The dwReturn member is set to the play speed in frames per second. The MCIPIONR.DRV device driver returns MCIERR_UNSUPPORTED_FUNCTION.
    '
    ''The following additional flags are used with the waveaudio device type. These constants are used in the dwItem member of the structure pointed to by the lpStatus parameter when MCI_STATUS_ITEM is specified for the dwFlags parameter.
    '
    'MCI_WAVE_FORMATTAG
    '
    '    ''The dwReturn member is set to the current format tag used for playing, recording, and saving.
    'MCI_WAVE_INPUT
    '
    '    ''The dwReturn member is set to the wave input device used for recording. If no device is in use and no device has been explicitly set, then the error return is MCIERR_WAVE_INPUTUNSPECIFIED.
    'MCI_WAVE_OUTPUT
    '
    '    ''The dwReturn member is set to the wave output device used for playing. If no device is in use and no device has been explicitly set, then the error return is MCIERR_WAVE_OUTPUTUNSPECIFIED.
    'MCI_WAVE_STATUS_AVGBYTESPERSEC
    '
    '    ''The dwReturn member is set to the current bytes per second used for playing, recording, and saving.
    'MCI_WAVE_STATUS_BITSPERSAMPLE
    '
    '    ''The dwReturn member is set to the current bits per sample used for playing, recording, and saving PCM formatted data.
    'MCI_WAVE_STATUS_BLOCKALIGN
    '
    '    ''The dwReturn member is set to the current block alignment used for playing, recording, and saving.
    'MCI_WAVE_STATUS_CHANNELS
    '
    '    ''The dwReturn member is set to the current channel count used for playing, recording, and saving.
    'MCI_WAVE_STATUS_LEVEL
    '
    '    ''The dwReturn member is set to the current record or playback level of PCM formatted data. The value is returned as an 8- or 16-bit value, depending on the sample size used. The right or mono channel level is returned in the low-order word. The left channel level is returned in the high-order word.
    'MCI_WAVE_STATUS_SAMPLESPERSEC
    '
        ''The dwReturn member is set to the current samples per second used for playing, recording, and saving.
End Enum

Public Type MCI_STATUS_PARMS
    dwCallback As Long 'The low-order word specifies a window handle used for the MCI_NOTIFY flag.
    dwReturn As Long 'Contains information on return.
    dwItem As Long 'Capability being queried.
    dwTrack As Long 'Length or number of tracks.
End Type


Public Function mciSendMessage(ByVal wDeviceID As Long, ByVal uMessage As MCI_COMMANDS, Optional ByVal dwParam1 As Long = 0, Optional ByVal dwParam2 As Long = 1, Optional ByVal ThrowExceptions As Boolean = True) As Long
    mciSendMessage = API_mciSendCommand(wDeviceID, uMessage, dwParam1, dwParam2)
    If ThrowExceptions Then _
        If mciSendMessage <> 0 Then _
            throw Exps.InvalidStatusException(mciGetErrorString(mciSendMessage))
End Function
Public Function mciSendString(ByVal StrCommand As String, Optional ByVal hWnd As Long = 0, Optional ByVal ThrowExceptions As Boolean = True) As String
    Dim Result As Long, Buffer As String
    Buffer = String$(MCI_COMMAND_LENGTH, 0)
    
    Result = API_mciSendString(StrCommand, ByVal StrPtr(Buffer), MCI_COMMAND_LENGTH, hWnd)
    If ThrowExceptions Then _
        If Result <> 0 Then _
            throw Exps.InvalidStatusException(mciGetErrorString(Result))
    
    If StrPtr(Buffer) <> 0 Then mciSendString = Mid$(Buffer, 1)
End Function
Public Function mciGetErrorString(ByVal dwError As Long) As String
    Dim Result As Long, Buffer As String
    Buffer = String$(MCI_COMMAND_LENGTH, 0)
    
    Result = API_mciGetErrorString(dwError, ByVal StrPtr(Buffer), MCI_COMMAND_LENGTH)
    If Result <> 0 Then throw Exps.SystemCallFailureException
    
    If StrPtr(Buffer) <> 0 Then mciGetErrorString = Mid$(Buffer, 1)
End Function


Public Function mciOpenFile(ByVal Path As String, Optional ByVal Alias As String, Optional ByVal Sharable As Boolean = True, Optional ByVal DeviceType As MCI_DEVICE_TYPES = MCI_DEVICE_TYPES.MCI_ALL_DEVICE_ID, Optional ByVal dwCallback As Long = 0, Optional ByVal dwDeviceFlags_MustBe0 As Long = 0)
    Dim OParam As MCI_OPEN_PARMS, Flags As Long
    OParam.dwCallback = dwCallback
    OParam.wDeviceID = 0
    If Sharable Then Flags = Flags Or MCI_OPEN_SHAREABLE
    If Alias <> "" Then
        OParam.lpstrAlias = Alias
        Flags = Flags Or MCI_OPEN_ALIAS
    End If
    If DeviceType <> MCI_ALL_DEVICE_ID Then
        OParam.lpstrDeviceType = DeviceType
        Flags = Flags Or MCI_OPEN_TYPE_ID Or MCI_OPEN_TYPE
    End If
    If dwDeviceFlags_MustBe0 <> 0 Then
        Call Memory.CopyMemory(VarPtr(OParam.lpstrElementName), VarPtr(dwDeviceFlags_MustBe0), 4)
        Flags = Flags Or MCI_OPEN_ELEMENT_ID
    ElseIf Path <> "" Then
        OParam.lpstrElementName = Path
        Flags = Flags Or MCI_OPEN_ELEMENT
    End If
    'open "path" alias ID
    Call mciSendMessage(0, MCI_OPEN, Flags, VarPtr(OParam))
    mciOpenFile = OParam.wDeviceID
End Function
Public Function mciPlayFile(ByVal ID As Long, ByVal IsLoop As Boolean, ByVal hWnd As Long) As Boolean
    
End Function
Public Function mciQueryInformation(ByVal ID As Long, ByVal Flags As MCI_STATUS_FLAGS, Optional ByVal dwItem As Long = 0, Optional ByVal dwCallback As Long = 0, Optional ByVal dwTrack As Long = 0, Optional ByVal ThrowExceptions As Boolean = True) As MCI_STATUS_PARMS
    mciQueryInformation.dwItem = dwItem
    mciQueryInformation.dwCallback = dwCallback
    mciQueryInformation.dwTrack = dwTrack
    Call mciSendMessage(ID, MCI_STATUS, Flags, VarPtr(mciQueryInformation), ThrowExceptions)
End Function
Public Function mciLength(ByVal ID As Long) As Long
    mciLength = mciQueryInformation(ID, MCI_STATUS_ITEM, MCI_STATUS_LENGTH).dwReturn
End Function
Public Function mciPosition(ByVal ID As Long) As Long
    mciPosition = mciQueryInformation(ID, MCI_STATUS_POSITION).dwReturn
End Function
Public Function mciMode(ByVal ID As Long) As MediaStreamState
    Dim Result As Long
    Result = mciQueryInformation(ID, MCI_STATUS_MODE, ThrowExceptions:=False).dwReturn
    Select Case Result
        Case 0: mciMode = mssNone
        Case MCI_MODES.MCI_MODE_OPEN: mciMode = mssOpen
        Case MCI_MODES.MCI_MODE_PLAY: mciMode = mssPlaying
        Case MCI_MODES.MCI_MODE_STOP: mciMode = mssStopped
        Case MCI_MODES.MCI_MODE_SEEK: mciMode = mssSeek
        Case MCI_MODES.MCI_MODE_PAUSE: mciMode = mssPaused
        Case Else: mciMode = mssCustom Or dt_mdMCIDevice Or Result
    End Select
End Function


Public Function mciQueryInformationStr(ByVal ID As String, ByVal Name As String, Optional ByVal ThrowExceptions As Boolean = True) As String
    mciQueryInformationStr = Mid$(mciSendString(MCI_FUNC_STATUS & " " & ID & " " & Name, 0, ThrowExceptions), 1)
End Function
Public Function mciLengthStr(ByVal ID As String) As Long
    MsgBox mciQueryInformationStr(ID, "length")
    mciLengthStr = Convert.ToLong(mciQueryInformationStr(ID, "length"), True, 1)
End Function
Public Function mciPositionStr(ByVal ID As String) As Long
    mciPositionStr = Convert.ToLong(mciQueryInformationStr(ID, "position"), True, 1)
End Function

Public Function mciIsPlayingStr(ByVal ID As String) As Boolean
    mciIsPlayingStr = (mciQueryInformationStr(ID, "mode", False) = "playing")
End Function

Public Sub mciOpenFileStr(ByVal Path As String, ByVal ID As String, ByVal hWnd As Long)
    'open "path" alias ID
    Call mciSendString(MCI_FUNC_OPEN & " """ & Path & """ alias " & ID, hWnd)
End Sub
Public Function mciPlayFileStr(ByVal ID As String, ByVal IsLoop As Boolean, ByVal hWnd As Long) As Boolean
    Dim Command As String
    Command = MCI_FUNC_PLAY & " " & ID
    If hWnd <> 0 Then Command = Command & " notify"
    If IsLoop Then Command = Command & " REPEAT"
    Call mciSendString(Command, hWnd)
    mciPlayFileStr = True
End Function
Public Sub mciExecuteFunction(ByVal ID As String, ByVal FuncName As String, Optional ByVal Flags As String = "", Optional ByVal hWnd As Long = 0, Optional ByVal ThrowExceptions As Boolean = True)
    Call mciSendString(LCase$(FuncName) & " " & ID & IIf(Flags = "", "", " " & Flags), hWnd, ThrowExceptions)
End Sub
Public Sub mciSetPositionStr(ByVal ID As String, ByVal Miliseconds As Variant)
    If mciIsPlayingStr(ID) Then
        Call mciExecuteFunction(ID, MCI_FUNC_PLAY, "from " & Miliseconds)
    Else
        Call mciExecuteFunction(ID, MCI_FUNC_SEEK, "to " & Miliseconds)
    End If
End Sub
Public Sub mciSeekMediaStr(ByVal ID As String, ByVal Origin As SeekOrigin, ByVal Miliseconds As Variant)
    Select Case Origin
        Case FromBeginning
            If Miliseconds = 0 Then
                If mciIsPlayingStr(ID) Then
                    Call mciExecuteFunction(ID, MCI_FUNC_PLAY, "from start")
                Else
                    Call mciExecuteFunction(ID, MCI_FUNC_SEEK, "to start")
                End If
            Else
                If mciIsPlayingStr(ID) Then
                    Call mciExecuteFunction(ID, MCI_FUNC_PLAY, "from " & Miliseconds)
                Else
                    Call mciExecuteFunction(ID, MCI_FUNC_SEEK, "to " & Miliseconds)
                End If
            End If
        Case FromEnd
            If Miliseconds = 0 Then
                If mciIsPlayingStr(ID) Then
                    Call mciExecuteFunction(ID, MCI_FUNC_PLAY, "from end")
                Else
                    Call mciExecuteFunction(ID, MCI_FUNC_SEEK, "to end")
                End If
            Else
                If mciIsPlayingStr(ID) Then
                    Call mciExecuteFunction(ID, MCI_FUNC_PLAY, "from " & (mciLengthStr(ID) - Miliseconds))
                Else
                    Call mciExecuteFunction(ID, MCI_FUNC_SEEK, "to " & (mciLengthStr(ID) - Miliseconds))
                End If
            End If
        Case FromCurrent
            If mciIsPlayingStr(ID) Then
                Call mciExecuteFunction(ID, MCI_FUNC_PLAY, "from " & (mciPositionStr(ID) + Miliseconds))
            Else
                Call mciExecuteFunction(ID, MCI_FUNC_SEEK, "to " & (mciPositionStr(ID) + Miliseconds))
            End If
    End Select
End Sub

