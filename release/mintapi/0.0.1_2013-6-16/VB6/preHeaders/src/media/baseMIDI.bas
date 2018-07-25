Attribute VB_Name = "baseMIDI"
Option Explicit
Const CLASSID As String = "MIDIAPI"
'--------------------------------------------------------------------
'--------------------------------------------------------------------
'--------------------------------------------------------------------
'midi in/out section
'--------------------------------------------------------------------
'--------------------------------------------------------------------
'--------------------------------------------------------------------
#If Not REMOVE_API_SECTION_MIDI_IN_OUT Then
    Public Const API_MMSYSERR_BASE = 0
    Public Const API_MMSYSERR_NOERROR = 0
    Public Const API_MMSYSERR_ERROR = (API_MMSYSERR_BASE + 1)
    Public Const API_MMSYSERR_BADDEVICEID = (API_MMSYSERR_BASE + 2)
    Public Const API_MMSYSERR_NOTENABLED = (API_MMSYSERR_BASE + 3)
    Public Const API_MMSYSERR_ALLOCATED = (API_MMSYSERR_BASE + 4)
    Public Const API_MMSYSERR_INVALHANDLE = (API_MMSYSERR_BASE + 5)
    Public Const API_MMSYSERR_NODRIVER = (API_MMSYSERR_BASE + 6)
    Public Const API_MMSYSERR_NOMEM = (API_MMSYSERR_BASE + 7)
    Public Const API_MMSYSERR_NOTSUPPORTED = (API_MMSYSERR_BASE + 8)
    Public Const API_MMSYSERR_BADERRNUM = (API_MMSYSERR_BASE + 9)
    Public Const API_MMSYSERR_INVALFLAG = (API_MMSYSERR_BASE + 10)
    Public Const API_MMSYSERR_INVALPARAM = (API_MMSYSERR_BASE + 11)
    Public Const API_MMSYSERR_HANDLEBUSY = (API_MMSYSERR_BASE + 12)
    Public Const API_MMSYSERR_INVALIDALIAS = (API_MMSYSERR_BASE + 13)
    Public Const API_MMSYSERR_LASTERROR = (API_MMSYSERR_BASE + 13)
    
    Public Const API_MM_MIM_CLOSE = &H3C2
    Public Const API_MM_MIM_DATA = &H3C3
    Public Const API_MM_MIM_ERROR = &H3C5
    Public Const API_MM_MIM_LONGDATA = &H3C4
    Public Const API_MM_MIM_LONGERROR = &H3C6
    Public Const API_MM_MIM_MOREDATA = &H3CC
    Public Const API_MM_MIM_OPEN = &H3C1
    
    Public Type API_MIDIHDR
        lpData As String
        dwBufferLength As Long
        dwBytesRecorded As Long
        dwUser As Long
        dwFlags As Long
        lpNext As Long
        Reserved As Long
    End Type
    Public Type API_MIDIINCAPS
        wMid As Integer
        wPid As Integer
        vDriverVersion As Long
        szPname As String * API_MAXPNAMELEN
    End Type
    Public Type API_MIDIOUTCAPS
       wMid As Integer                   ' Manufacturer identifier of the device driver for the MIDI output device
                                         ' For a list of identifiers, see the Manufacturer Indentifier topic in the
                                         ' Multimedia Reference of the Platform SDK.
       
       wPid As Integer                   ' Product Identifier Product of the MIDI output device. For a list of
                                         ' product identifiers, see the Product Identifiers topic in the Multimedia
                                         ' Reference of the Platform SDK.
       
       vDriverVersion As Long            ' Version number of the device driver for the MIDI output device.
                                         ' The high-order byte is the major version number, and the low-order byte is
                                         ' the minor version number.
                                         
       szPname As String * API_MAXPNAMELEN   ' Product name in a null-terminated string.
       
       wTechnology As Integer            ' One of the following that describes the MIDI output device:
                                         '     MOD_FMSYNTH-The device is an FM synthesizer.
                                         '     MOD_MAPPER-The device is the Microsoft MIDI mapper.
                                         '     MOD_MIDIPORT-The device is a MIDI hardware port.
                                         '     MOD_SQSYNTH-The device is a square wave synthesizer.
                                         '     MOD_SYNTH-The device is a synthesizer.
                                         
       wVoices As Integer                ' Number of voices supported by an internal synthesizer device. If the
                                         ' device is a port, this member is not meaningful and is set to 0.
                                         
       wNotes As Integer                 ' Maximum number of simultaneous notes that can be played by an internal
                                         ' synthesizer device. If the device is a port, this member is not meaningful
                                         ' and is set to 0.
                                         
       wChannelMask As Integer           ' Channels that an internal synthesizer device responds to, where the least
                                         ' significant bit refers to channel 0 and the most significant bit to channel
                                         ' 15. Port devices that transmit on all channels set this member to 0xFFFF.
                                         
       dwSupport As Long                 ' One of the following describes the optional functionality supported by
                                         ' the device:
                                         '     MIDICAPS_CACHE-Supports patch caching.
                                         '     MIDICAPS_LRVOLUME-Supports separate left and right volume control.
                                         '     MIDICAPS_STREAM-Provides direct support for the midiStreamOut function.
                                         '     MIDICAPS_VOLUME-Supports volume control.
                                         '
                                         ' If a device supports volume changes, the MIDICAPS_VOLUME flag will be set
                                         ' for the dwSupport member. If a device supports separate volume changes on
                                         ' the left and right channels, both the MIDICAPS_VOLUME and the
                                         ' MIDICAPS_LRVOLUME flags will be set for this member.
    End Type
    Public Type API_MIDIPROPTEMPO
        cbStruct As Long
        dwTempo As Long
    End Type
    Public Type API_MIDIPROPTIMEDIV
        cbStruct As Long
        dwTimeDiv As Long
    End Type
    Public Type API_MIDISTRMBUFFVER
        dwVersion As Long                  '  Stream buffer format version
        dwMid As Long                      '  Manufacturer ID as defined in MMREG.H
        dwOEMVersion As Long               '  Manufacturer version for custom ext
    End Type

    Public Declare Function API_midiConnect Lib "winmm" Alias "midiConnect" (ByVal hmi As Long, ByVal hmo As Long, pReserved As Any) As Long
    Public Declare Function API_midiDisconnect Lib "winmm" Alias "midiDisconnect" (ByVal hmi As Long, ByVal hmo As Long, pReserved As Any) As Long
    Public Declare Function API_midiInAddBuffer Lib "winmm" Alias "midiInAddBuffer" (ByVal hMidiIn As Long, lpMidiInHdr As API_MIDIHDR, ByVal uSize As Long) As Long
    Public Declare Function API_midiInClose Lib "winmm" Alias "midiInClose" (ByVal hMidiIn As Long) As Long
    Public Declare Function API_midiInGetDevCaps Lib "winmm" Alias "midiInGetDevCapsA" (ByVal uDeviceID As Long, lpCaps As API_MIDIINCAPS, ByVal uSize As Long) As Long
    Public Declare Function API_midiInGetErrorText Lib "winmm" Alias "midiInGetErrorTextA" (ByVal err As Long, ByVal lpText As String, ByVal uSize As Long) As Long
    Public Declare Function API_midiInGetID Lib "winmm" Alias "midiInGetID" (ByVal hMidiIn As Long, lpuDeviceID As Long) As Long
    Public Declare Function API_midiInGetNumDevs Lib "winmm" Alias "midiInGetNumDevs" () As Long
    Public Declare Function API_midiInMessage Lib "winmm" Alias "midiInMessage" (ByVal hMidiIn As Long, ByVal msg As Long, ByVal dw1 As Long, ByVal dw2 As Long) As Long
    Public Declare Function API_midiInOpen Lib "winmm" Alias "midiInOpen" (lphMidiIn As Long, ByVal uDeviceID As Long, ByVal dwCallback As Long, ByVal dwInstance As Long, ByVal dwFlags As Long) As Long
    Public Declare Function API_midiInPrepareHeader Lib "winmm" Alias "midiInPrepareHeader" (ByVal hMidiIn As Long, lpMidiInHdr As API_MIDIHDR, ByVal uSize As Long) As Long
    Public Declare Function API_midiInReset Lib "winmm" Alias "midiInReset" (ByVal hMidiIn As Long) As Long
    Public Declare Function API_midiInStart Lib "winmm" Alias "midiInStart" (ByVal hMidiIn As Long) As Long
    Public Declare Function API_midiInStop Lib "winmm" Alias "midiInStop" (ByVal hMidiIn As Long) As Long
    Public Declare Function API_midiInUnprepareHeader Lib "winmm" Alias "midiInUnprepareHeader" (ByVal hMidiIn As Long, lpMidiInHdr As API_MIDIHDR, ByVal uSize As Long) As Long
    Public Declare Function API_midiOutCacheDrumPatches Lib "winmm" Alias "midiOutCacheDrumPatches" (ByVal hMidiOut As Long, ByVal uPatch As Long, lpKeyArray As Long, ByVal uFlags As Long) As Long
    Public Declare Function API_midiOutCachePatches Lib "winmm" Alias "midiOutCachePatches" (ByVal hMidiOut As Long, ByVal uBank As Long, lpPatchArray As Long, ByVal uFlags As Long) As Long
    Public Declare Function API_midiOutClose Lib "winmm" Alias "midiOutClose" (ByVal hMidiOut As Long) As Long
    Public Declare Function API_midiOutGetDevCaps Lib "winmm" Alias "midiOutGetDevCapsA" (ByVal uDeviceID As Long, lpCaps As API_MIDIOUTCAPS, ByVal uSize As Long) As Long
    Public Declare Function API_midiOutGetErrorText Lib "winmm" Alias "midiOutGetErrorTextA" (ByVal err As Long, ByVal lpText As String, ByVal uSize As Long) As Long
    Public Declare Function API_midiOutGetID Lib "winmm" Alias "midiOutGetID" (ByVal hMidiOut As Long, lpuDeviceID As Long) As Long
    Public Declare Function API_midiOutGetNumDevs Lib "winmm" Alias "midiOutGetNumDevs" () As Integer
    Public Declare Function API_midiOutGetVolume Lib "winmm" Alias "midiOutGetVolume" (ByVal uDeviceID As Long, lpdwVolume As Long) As Long
    Public Declare Function API_midiOutLongMsg Lib "winmm" Alias "midiOutLongMsg" (ByVal hMidiOut As Long, lpMidiOutHdr As API_MIDIHDR, ByVal uSize As Long) As Long
    Public Declare Function API_midiOutOpen Lib "winmm" Alias "midiOutOpen" (lphMidiOut As Long, ByVal uDeviceID As Long, ByVal dwCallback As Long, ByVal dwInstance As Long, ByVal dwFlags As Long) As Long
    Public Declare Function API_midiOutMessage Lib "winmm" Alias "midiOutMessage" (ByVal hMidiOut As Long, ByVal msg As Long, ByVal dw1 As Long, ByVal dw2 As Long) As Long
    Public Declare Function API_midiOutPrepareHeader Lib "winmm" Alias "midiOutPrepareHeader" (ByVal hMidiOut As Long, lpMidiOutHdr As API_MIDIHDR, ByVal uSize As Long) As Long
    Public Declare Function API_midiOutReset Lib "winmm" Alias "midiOutReset" (ByVal hMidiOut As Long) As Long
    Public Declare Function API_midiOutSetVolume Lib "winmm" Alias "midiOutSetVolume" (ByVal uDeviceID As Long, ByVal dwVolume As Long) As Long
    Public Declare Function API_midiOutShortMsg Lib "winmm" Alias "midiOutShortMsg" (ByVal hMidiOut As Long, ByVal dwMsg As Long) As Long
    Public Declare Function API_midiOutUnprepareHeader Lib "winmm" Alias "midiOutUnprepareHeader" (ByVal hMidiOut As Long, lpMidiOutHdr As API_MIDIHDR, ByVal uSize As Long) As Long
    Public Declare Function API_midiStreamClose Lib "winmm" Alias "midiStreamClose" (ByVal hms As Long) As Long
    Public Declare Function API_midiStreamOpen Lib "winmm" Alias "midiStreamOpen" (phms As Long, puDeviceID As Long, ByVal cMidi As Long, ByVal dwCallback As Long, ByVal dwInstance As Long, ByVal fdwOpen As Long) As Long
    Public Declare Function API_midiStreamOut Lib "winmm" Alias "midiStreamOut" (ByVal hms As Long, pmh As API_MIDIHDR, ByVal cbmh As Long) As Long
    Public Declare Function API_midiStreamPause Lib "winmm" Alias "midiStreamPause" (ByVal hms As Long) As Long
    Public Declare Function API_midiStreamPosition Lib "winmm" Alias "midiStreamPosition" (ByVal hms As Long, lpmmt As API_MMTIME, ByVal cbmmt As Long) As Long
    Public Declare Function API_midiStreamRestart Lib "winmm" Alias "midiStreamRestart" (ByVal hms As Long) As Long
    Public Declare Function API_midiStreamProperty Lib "winmm" Alias "midiStreamProperty" (ByVal hms As Long, lppropdata As Byte, ByVal dwProperty As Long) As Long
    Public Declare Function API_midiStreamStop Lib "winmm" Alias "midiStreamStop" (ByVal hms As Long) As Long
#End If

'Preface
'After the MIDI 1.0 standard was finalized in the early 1980's, numerous musical instruments with MIDI jacks appeared upon the market. Musicians started to attach these instruments via their MIDI ports, and quickly discovered that the MIDI 1.0 specification had overlooked some important concerns.
'One typical scenario may have been as follows:
'A musician attaches his Roland D-10 to his Yamaha DX-7, because he prefers the front panel of the D-10, but prefers the sound of the DX-7, and he wants to use the D-10 to "play" the DX-7. He selects the patch labeled "Piano" on the D-10, and he plays the D-10 keyboard, and on the DX-7 he hears... a trumpet? How did this happen? Well, it happened because MIDI sends a program change message that contains only a patch number -- not the actual name of the patch. So if patch #1 on the DX-7 is a trumpet sound, then that's what he gets on the DX-7, despite the fact that selecting patch #1 on the D-10 yields a piano sound on the D-10. The MIDI 1.0 specification did not require that particular sounds be assigned to particular patch numbers, so every manufacturer used his own discretion as to "patch mapping".
'But the real problem was with MIDI files that the musician made. MIDI files contain only MIDI messages. So, any program change event in a MIDI file refers only to a patch number as well -- not the actual patch name. So, this musician creates a MIDI file using his D-10. He has a piano track, so he puts a program change event at the track's beginning to select patch #1, which happens to be a Piano sound on his D-10. He takes that MIDI file to a friend's house. The friend has a DX-7. They play the MIDI file on that DX-7, and suddenly, the piano part is playing with a trumpet sound. Well, that's because patch #1 on the DX-7 is not a piano -- it's a trumpet sound. To "fix" the MIDI file, now the musician with the DX-7 has to edit the MIDI tracks and change every MIDI Program Change event so that it refers to the correct patch number on his DX-7. This deviance among MIDI sound modules made it very difficult for musicians to create MIDI arrangements that played properly upon various MIDI sound modules.
'There were also some other deviances among early MIDI modules that made it more difficult to use them together via MIDI. To address these concerns, Roland proposed an addendum to the MIDI 1.0 specification in the late 1980's. This new addendum was called "General MIDI" (GM). It added some new requirements to the base MIDI 1.0 specification (but does not supplant any parts of the 1.0 specification -- the 1.0 specification is still the base level to which all MIDI devices should adhere). GM has now been adopted as part of the MIDI 2.0 specification.
'
'
'--------------------------------------------------------------------------------
'
'
'General MIDI Patches
'So to make MIDI Program Change messages of more practical use, Roland found it necessary to adopt a standard "patch bank". In other words, what was needed was to assign specific instrument sounds to specific patch numbers. For example, it was decided that patch number 1 upon all sound modules should be the sound of an Acoustic Grand Piano. In this way, no matter what MIDI sound module you use, when you select patch number 1, you always hear some sort of Acoustic Grand Piano sound. A standard was set for 128 patches which must appear in a specific order, and this standard is called General MIDI (GM). For example, patch number 25 upon a GM module must be a Nylon String Guitar. The chart, GM Patches, shows you the names of all GM Patches, and their respective Program Change numbers.
'Nowadays, most modules (including the built-in sound modules of computer sound cards) ship with a GM bank (of 128 patches) so that it is easy to play MIDI files upon any MIDI module, without needing to edit all of the Program Change events in the file.
'
'
'--------------------------------------------------------------------------------
'
'
'General MIDI Multi-Timbral requirement
'Another burgeoning technology in the late 1980's was the multi-timbral module. Typically, there were deviances in the way that various manufacturers implemented this, since the 1.0 specification did not specifically address such devices. For example, some early multi-timbral modules supported only a limited set of the 16 MIDI channels simultaneously, so if you had a MIDI file with tracks upon unsupported MIDI channels, you wouldn't hear those tracks play back. You may not have even realized that those parts weren't being played.
'So, one requirement of a GM-compliant module is that it must be fully multi-timbral, meaning that it can play MIDI messages upon all 16 channels simultaneously, with a different GM Patch sounding for each channel.
'
'
'--------------------------------------------------------------------------------
'
'
'General MIDI Note Number assignments
'There were also deviances in regards to Note Number mapping. For example, some manufacturers mapped middle 'C' to MIDI Note Number 60. Others mapped it to Note Numbers 72 or 48. Some modules even had middle C mapped to various places in different patches, depending upon the instrument. For example, a bass guitar patch may have middle C mapped to the highest C on the keyboard (since the most useful range on a bass guitar is below middle C). A flute patch may have middle C mapped to the lowest C on the keyboard.
'The result was that, it became confusing to keep track of which key (ie, MIDI Note Number) played middle C for each patch. Also, when a MIDI track was played back upon certain modules, the part may play back an octave too high or low.
'It therefore was decided that all patches must sound an A440 pitch when receiving a MIDI note number of 69. (ie, Note Number 69 plays the A above middle C, and therefore Note Number 60 is middle C).
'There were deviances in regards to "drum machines" as well. Most drum machines (and drum units built into multi-timbral modules) play a different drum sound for each MIDI Note Number. But the 1.0 specification never spelled out which drum sounds were assigned to which MIDI note numbers. So, whereas note number 60 may play a snare upon one drum unit, upon another drum unit, it may play a crash cymbal. Again, this caused trouble with MIDI files, since sometimes a drum part would play back with the wrong drum sounds.
'To address this discrepancy, the GM addendum contains a "drum map". This assigns about 48 common drum sounds to 48 specific MIDI Note Numbers. The assignments of drum sounds to MIDI notes is shown in the chart, GM Drum Sounds. Also, it was decided that a GM drum unit should default to using MIDI channel 10 to receive MIDI messages. Therefore, a composer of a GM MIDI file can safely assume that his drum part will play correctly if he uses the GM Drum note assignments and records the drum part upon MIDI channel 10.
'
'
'--------------------------------------------------------------------------------
'
'
'General MIDI polyphony
'Polyphony is how many notes a module can sound simultaneously. For example, perhaps a module can sound 32 notes simultaneously. Early MIDI modules typically had very limited polyphony. For example, the Prophet 5 could sound only 5 notes simultaneously.
'This discrepancy in polyphony among MIDI modules made it difficult for arrangers to create MIDI files that played properly upon various modules. For example, if the arranger created too "busy" an arrangement, it could exceed the polyphony of a particular module, and therefore some of the notes may not be heard.
'To address this discrepancy, the GM addendum stipulated that a GM module should be capable of sounding at least 24 notes simultaneously. (Ie, It must have 24 note polyphony). It could exceed 24 note polyphony, but it had to have at least that level of polyphony. In this way, if an arranger ensured that he never had more than 24 notes sounding simultaneously in his MIDI file, all notes of his arrangement would be heard upon any GM module.
'
'
'--------------------------------------------------------------------------------
'
'
'Other General MIDI requirements
'Finally, the GM addendum attempted to address some other discrepancies by spelling out a few more requirements.
'A GM module should respond to velocity (ie, for note messages). This typically controls the VCA level (ie, volume) of each note, but the GM addendum unfortunately did not set a specific function for velocity. Some modules may allow velocity to affect other parameters on some patches.
'The pitch wheel bend range should default to +/- 2 semitones. This allows an arranger to use pitch bend messages in his arrangement without worrying whether a bend that is supposed to be up 2 whole steps will instead jump up 2 octaves upon a certain sound module.
'The module also should respond to Channel Pressure (often used to control VCA level or VCO level for vibrato depth). Again, the GM addendum unfortunately did not set a specific function for channel pressure, although typically it defaults to controlling the volume of a note while it is being held.
'Finally, a GM module should also respond to the following MIDI controller messages: Modulation (1) (usually hard-wired to control LFO amount, ie, vibrato), Channel Volume (7), Pan (10), Expression (11), Sustain (64), Reset All Controllers (121), and All Notes Off (123). Additionally, the module should respond to these Registered Parameter Numbers: Pitch Wheel Bend Range (0), Fine Tuning (1), and Coarse Tuning (2).
'There were also some default settings that a GM module should apply upon power up. Channel Volume should default to 90, with all other controllers and effects off (including pitch wheel offset of 0). Initial tuning should be standard, A440 reference.
'
'
'--------------------------------------------------------------------------------
'
'
'General MIDI messages
'The GM addendum did specify a couple System Exclusive messages to alter settings that are common to all GM units, but which were not addressed by the 1.0 specification.
'One such message is for Master Volume -- not just the volume of a patch upon any one MIDI channel, but the master volume of the module itself.
'There is also a System Exclusive message that can be used to turn a module's General MIDI mode on or off. This is useful for modules that also offer more expansive, non-GM playback modes or extra, programmable banks of patches beyond the GM set, but need to allow the musician to switch to GM mode when desired.
'
'
'--------------------------------------------------------------------------------
'
'
'Conclusion
'GM Standard makes it easy for musicians to put Program Change messages in their MIDI (sequencer) song files, confident that those messages will select the correct instruments on all GM sound modules, and the song file would therefore play all of the correct instrumentation automatically. Furthermore, musicians need not worry about parts being played back in the wrong octave. Finally, musicians didn't have to worry that a snare drum part, for example, would be played back on a Cymbal. The GM Standard also spells out other minimum requirements that a GM module should meet, such as being able to respond to Pitch and Modulation Wheels, and also being able to play 24 notes simultaneously (with dynamic voice allocation between the 16 Parts). All of these standards help to ensure that MIDI Files play back properly upon setups of various equipment.
'The GM standard is actually not encompassed in the MIDI specification proper (ie, it's an addendum), and there's no reason why someone can't set up the Patches in his sound module to be entirely different sounds than the GM set. After all, most MIDI sound modules offer such programmability. But, most have a GM option so that musicians can easily play the many MIDI files that expect a GM module.
'NOTE: The GM Standard doesn't dictate how a module produces sound. For example, one module could use cheap FM synthesis to simulate the Acoustic Grand Piano patch. Another module could use 24 digital audio waveforms of various notes on a piano, mapped out across the MIDI note range, to create that one Piano patch. Obviously, the 2 patches won't sound exactly alike, but at least they will both be piano patches on the 2 modules. So too, GM doesn't dictate VCA envelopes for the various patches, so for example, the Sax patch upon one module may have a longer release time than the same patch upon another module.
'
'
'--------------------------------------------------------------------------------



'GM patches
'This chart shows the names of all 128 GM Instruments, and the MIDI Program Change numbers which select those Instruments.
'The patches are arranged into 16 "families" of instruments, with each family containing 8 instruments. For example, there is a Reed family. Among the 8 instruments within the Reed family, you will find Saxophone, Oboe, and Clarinet.

'Prog#   Instrument            Prog#    Instrument
'
' PIANO                           CHROMATIC PERCUSSION
'1    Acoustic Grand             9   Celesta
'2    Bright Acoustic           10   Glockenspiel
'3    Electric Grand            11   Music Box
'4    Honky-Tonk                12   Vibraphone
'5    Electric Piano 1          13   Marimba
'6    Electric Piano 2          14   Xylophone
'7    Harpsichord               15   Tubular Bells
'8    Clavinet                  16   Dulcimer
'
'  Organ GUITAR
'17   Drawbar Organ             25   Nylon String Guitar
'18   Percussive Organ          26   Steel String Guitar
'19   Rock Organ                27   Electric Jazz Guitar
'20   Church Organ              28   Electric Clean Guitar
'21   Reed Organ                29   Electric Muted Guitar
'22   Accoridan                 30   Overdriven Guitar
'23   Harmonica                 31   Distortion Guitar
'24   Tango Accordian           32   Guitar Harmonics
'
'  BASS                           SOLO STRINGS
'33   Acoustic Bass             41   Violin
'34   Electric Bass(finger)     42   Viola
'35   Electric Bass(pick)       43   Cello
'36   Fretless Bass             44   Contrabass
'37   Slap Bass 1               45   Tremolo Strings
'38   Slap Bass 2               46   Pizzicato Strings
'39   Synth Bass 1              47   Orchestral Strings
'40   Synth Bass 2              48   Timpani
'
'  Ensemble                      BRASS
'49   String Ensemble 1         57   Trumpet
'50   String Ensemble 2         58   Trombone
'51   SynthStrings 1            59   Tuba
'52   SynthStrings 2            60   Muted Trumpet
'53   Choir Aahs                61   French Horn
'54   Voice Oohs                62   Brass Section
'55   Synth Voice               63   SynthBrass 1
'56   Orchestra Hit             64   SynthBrass 2
'
'  REED                         PIPE
'65   Soprano Sax               73   Piccolo
'66   Alto Sax                  74   Flute
'67   Tenor Sax                 75   Recorder
'68   Baritone Sax              76   Pan Flute
'69   Oboe                      77   Blown Bottle
'70   English Horn              78   Skakuhachi
'71   Bassoon                   79   Whistle
'72   Clarinet                  80   Ocarina
'
'  SYNTH LEAD                     SYNTH PAD
'81   Lead 1 (square)           89   Pad 1 (new age)
'82   Lead 2 (sawtooth)         90   Pad 2 (warm)
'83   Lead 3 (calliope)         91   Pad 3 (polysynth)
'84   Lead 4 (chiff)            92   Pad 4 (choir)
'85   Lead 5 (charang)          93   Pad 5 (bowed)
'86   Lead 6 (voice)            94   Pad 6 (metallic)
'87   Lead 7 (fifths)           95   Pad 7 (halo)
'88   Lead 8 (bass+lead)        96   Pad 8 (sweep)
'
'   SYNTH EFFECTS                  ETHNIC
' 97  FX 1 (rain)               105   Sitar
' 98  FX 2 (soundtrack)         106   Banjo
' 99  FX 3 (crystal)            107   Shamisen
'100  FX 4 (atmosphere)         108   Koto
'101  FX 5 (brightness)         109   Kalimba
'102  FX 6 (goblins)            110   Bagpipe
'103  FX 7 (echoes)             111   Fiddle
'104  FX 8 (sci-fi)             112   Shanai
'
'   PERCUSSIVE                     SOUND EFFECTS
'113  Tinkle Bell               121   Guitar Fret Noise
'114  Agogo                     122   Breath Noise
'115  Steel Drums               123   Seashore
'116  Woodblock                 124   Bird Tweet
'117  Taiko Drum                125   Telephone Ring
'118  Melodic Tom               126   Helicopter
'119  Synth Drum                127   Applause
'120  Reverse Cymbal            128   Gunshot

'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>CONST LIST 1

'Prog# refers to the MIDI Program Change number that causes this Patch to be selected. These decimal numbers are what the user normally sees on his module's display (or in a sequencer's "Event List"), but note that MIDI modules count the first Patch as 0, not 1. So, the value that is sent in the Program Change message would actually be one less. For example, the Patch number for Reverse Cymbal is actually sent as 119 rather than 120. But, when entering that Patch number using sequencer software or your module's control panel, the software or module understands that humans normally start counting from 1, and so would expect that you'd count the Reverse Cymbal as Patch 120. Therefore, the software or module automatically does this subtraction when it generates the MIDI Program Change message.
'
'So, sending a MIDI Program Change with a value of 120 (ie, actually 119) to a Part causes the Reverse Cymbal Patch to be selected for playing that Part's MIDI data.



'GM Drum Sounds
'This chart shows what drum sounds are assigned to each MIDI note for a GM module (ie, that has a drum part).
'
'
'MIDI    Drum Sound          MIDI    Drum Sound
'Note #                      Note #
' 35   Acoustic Bass Drum     59   Ride Cymbal 2
' 36   Bass Drum 1            60   Hi Bongo
' 37   Side Stick             61   Low Bongo
' 38   Acoustic Snare         62   Mute Hi Conga
' 39   Hand Clap              63   Open Hi Conga
' 40   Electric Snare         64   Low Conga
' 41   Low Floor Tom          65   High Timbale
' 42   Closed Hi-Hat          66   Low Timbale
' 43   High Floor Tom         67   High Agogo
' 44   Pedal Hi-Hat           68   Low Agogo
' 45   Low Tom                69   Cabasa
' 46   Open Hi-Hat            70   Maracas
' 47   Low-Mid Tom            71   Short Whistle
' 48   Hi-Mid Tom             72   Long Whistle
' 49   Crash Cymbal 1         73   Short Guiro
' 50   High Tom               74   Long Guiro
' 51   Ride Cymbal 1          75   Claves
' 52   Chinese Cymbal         76   Hi Wood Block
' 53   Ride Bell              77   Low Wood Block
' 54   Tambourine             78   Mute Cuica
' 55   Splash Cymbal          79   Open Cuica
' 56   Cowbell                80   Mute Triangle
' 57   Crash Cymbal 2         81   Open Triangle
'58    Vibraslap
'
'A note-on with note number 42 will trigger a Closed Hi-Hat. This should cut off any Open Hi-Hat or Pedal Hi-Hat sound that may be sustaining. So too, a Pedal Hi-Hat should cut off a sustaining Open Hi-Hat or Closed Hi-Hat. In other words, only one of these three drum sounds can be sounding at any given time.
'Similiarly, a Short Whistle should cut off a Long Whistle. A Short Guiro should cut off a Long Guiro. An Mute Triangle should cut off an Open Triangle. A Mute Cuica should cut off an Open Cuica.
'
'Normally, all the above drum sounds have a fixed duration. Regardless of the time between when a Note-On is received and when a matching Note-Off is received, the drum sound always plays for a given duration. For example, assume that a device has a "Crash Cymbal 1" sound that plays for 4 seconds. If a Note-On for note number 49 is received, that cymbal sound starts playing. If a Note-Off for note number 49 is received only 1 second later, that should not cut off the remaining 3 seconds of the sound. The exceptions may be Long Whistle and Long Guiro, which may use the duration between the Note-On and Note-off to determine how "long" the sound plays.
'
'If a drum is still sounding when another one of its Note-Ons is received, typically, another voice "stacks" another instance of that sound playing.


'------------------------------------------------------------------------------------------------------------------------'


'------------------------------------------------------------'
'------------------------------------------------------------'
'---------------|         Const Lists        |---------------'
'------------------------------------------------------------'
'------------------------------------------------------------'


'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<CONST LIST 1
'***** GM INSTRUMENTS *****

'' Piano     - Pianos
Public Const GM_INSTRUMENT_Piano__Acoustic_Grand = 1 '1    Acoustic Grand
Public Const GM_INSTRUMENT_Piano__Bright_Acoustic = 2 '2    Bright Acoustic
Public Const GM_INSTRUMENT_Piano__Electric_Grand = 3 '3    Electric Grand
Public Const GM_INSTRUMENT_Piano__Honky_Tonk = 4 '4    Honky -Tonk
Public Const GM_INSTRUMENT_Piano__Electric_Piano_1 = 5 '5    Electric Piano 1
Public Const GM_INSTRUMENT_Piano__Electric_Piano_2 = 6 '6    Electric Piano 2
Public Const GM_INSTRUMENT_Piano__Harpsichord = 7 '7    Harpsichord
Public Const GM_INSTRUMENT_Piano__Clavinet = 8 '8    Clavinet
Public Const GM_INSTRUMENT_Piano__Default = GM_INSTRUMENT_Piano__Acoustic_Grand

'' Chromatic Percussion     - Tuned Idiophones
Public Const GM_INSTRUMENT_Chromatic_Percussion__Celesta = 9 '9     Celesta
Public Const GM_INSTRUMENT_Chromatic_Percussion__Glockenspiel = 10 '10     Glockenspiel
Public Const GM_INSTRUMENT_Chromatic_Percussion__Music_Box = 11 '11   Music Box
Public Const GM_INSTRUMENT_Chromatic_Percussion__Vibraphone = 12 '12   Vibraphone
Public Const GM_INSTRUMENT_Chromatic_Percussion__Marimba = 13 '13   Marimba
Public Const GM_INSTRUMENT_Chromatic_Percussion__Xylophone = 14 '14   Xylophone
Public Const GM_INSTRUMENT_Chromatic_Percussion__Tubular_Bells = 15 '15   Tubular Bells
Public Const GM_INSTRUMENT_Chromatic_Percussion__Dulcimer = 16 '16   Dulcimer
Public Const GM_INSTRUMENT_Chromatic_Percussion__Default = GM_INSTRUMENT_Chromatic_Percussion__Celesta

'' Organ     - Organs
Public Const GM_INSTRUMENT_Organ__Drawbar_Organ = 17 '17   Drawbar Organ
Public Const GM_INSTRUMENT_Organ__Percussive_Organ = 18 '18   Percussive Organ
Public Const GM_INSTRUMENT_Organ__Rock_Organ = 19 '19   Rock Organ
Public Const GM_INSTRUMENT_Organ__Church_Organ = 20 '20   Church Organ
Public Const GM_INSTRUMENT_Organ__Reed_Organ = 21 '21   Reed Organ
Public Const GM_INSTRUMENT_Organ__Accoridan = 22 '22   Accoridan
Public Const GM_INSTRUMENT_Organ__Harmonica = 23 '23   Harmonica
Public Const GM_INSTRUMENT_Organ__Tango_Accordian = 24 '24   Tango Accordian
Public Const GM_INSTRUMENT_Organ__Default = GM_INSTRUMENT_Organ__Drawbar_Organ


'' Guitar     - Guitars
Public Const GM_INSTRUMENT_Guitar__Nylon_String_Guitar = 25 '25   Nylon String Guitar
Public Const GM_INSTRUMENT_Guitar__Steel_String_Guitar = 26 '26   Steel String Guitar
Public Const GM_INSTRUMENT_Guitar__Electric_Jazz_Guitar = 27 '27   Electric Jazz Guitar
Public Const GM_INSTRUMENT_Guitar__Electric_Clean_Guitar = 28 '28   Electric Clean Guitar
Public Const GM_INSTRUMENT_Guitar__Electric_Muted_Guitar = 29 '29   Electric Muted Guitar
Public Const GM_INSTRUMENT_Guitar__Overdriven_Guitar = 30 '30      Overdriven Guitar
Public Const GM_INSTRUMENT_Guitar__Distortion_Guitar = 31 '31      Distortion Guitar
Public Const GM_INSTRUMENT_Guitar__Guitar_Harmonics = 32 '32      Guitar Harmonics
Public Const GM_INSTRUMENT_Guitar__Default = GM_INSTRUMENT_Guitar__Nylon_String_Guitar

'' Bass     - Basses
Public Const GM_INSTRUMENT_Bass__Acoustic_Bass = 33 '33   Acoustic Bass
Public Const GM_INSTRUMENT_Bass__Electric_Bass___finger = 34 '34   Electric Bass(finger)
Public Const GM_INSTRUMENT_Bass__Electric_Bass___pick = 35 '35   Electric Bass(pick)
Public Const GM_INSTRUMENT_Bass__Fretless_Bass = 36 '36   Fretless Bass
Public Const GM_INSTRUMENT_Bass__Slap_Bass_1 = 37 '37   Slap Bass 1
Public Const GM_INSTRUMENT_Bass__Slap_Bass_2 = 38 '38   Slap Bass 2
Public Const GM_INSTRUMENT_Bass__Synth_Bass_1 = 39 '39   Synth Bass 1
Public Const GM_INSTRUMENT_Bass__Synth_Bass_2 = 40 '40   Synth Bass 2
Public Const GM_INSTRUMENT_Bass__Default = GM_INSTRUMENT_Bass__Electric_Bass___finger

'' Solo Strings     - String And Timpani
Public Const GM_INSTRUMENT_Solo_Strings__Violin = 41 '41    Violin
Public Const GM_INSTRUMENT_Solo_Strings__Viola = 42 '42    Viola
Public Const GM_INSTRUMENT_Solo_Strings__Cello = 43 '43    Cello
Public Const GM_INSTRUMENT_Solo_Strings__Contrabass = 44 '44    Contrabass
Public Const GM_INSTRUMENT_Solo_Strings__Tremolo_Strings = 45 '45    Tremolo Strings
Public Const GM_INSTRUMENT_Solo_Strings__Pizzicato_Strings = 46 '46    Pizzicato Strings
Public Const GM_INSTRUMENT_Solo_Strings__Orchestral_Strings = 47 '47    Orchestral Strings
Public Const GM_INSTRUMENT_Solo_Strings__Timpani = 48 '48    Timpani
Public Const GM_INSTRUMENT_Solo_Strings__Default = GM_INSTRUMENT_Solo_Strings__Violin

'' Ensemble Brass     - Ensemble Strings And Voices
Public Const GM_INSTRUMENT_Ensemble__String_Ensemble_1 = 49 '49   String Ensemble 1
Public Const GM_INSTRUMENT_Ensemble__String_Ensemble_2 = 50 '50   String Ensemble 2
Public Const GM_INSTRUMENT_Ensemble__SynthStrings_1 = 51 '51   SynthStrings 1
Public Const GM_INSTRUMENT_Ensemble__SynthStrings_2 = 52 '52   SynthStrings 2
Public Const GM_INSTRUMENT_Ensemble__Choir_Aahs = 53 '53   Choir Aahs
Public Const GM_INSTRUMENT_Ensemble__Voice_Oohs = 54 '54   Voice Oohs
Public Const GM_INSTRUMENT_Ensemble__Synth_Voice = 55 '55   Synth Voice
Public Const GM_INSTRUMENT_Ensemble__Orchestra_Hit = 56 '56   Orchestra Hit
Public Const GM_INSTRUMENT_Ensemble__Default = GM_INSTRUMENT_Ensemble__String_Ensemble_1

'' Brass     - Brasses
Public Const GM_INSTRUMENT_Brass__Trumpet = 57 '57   Trumpet
Public Const GM_INSTRUMENT_Brass__Trombone = 58 '58   Trombone
Public Const GM_INSTRUMENT_Brass__Tuba = 59 '59   Tuba
Public Const GM_INSTRUMENT_Brass__Muted_Trumpet = 60 '60   Muted Trumpet
Public Const GM_INSTRUMENT_Brass__French_Horn = 61 '61   French Horn
Public Const GM_INSTRUMENT_Brass__Brass_Section = 62 '62   Brass Section
Public Const GM_INSTRUMENT_Brass__SynthBrass_1 = 63 '63    SynthBrass 1
Public Const GM_INSTRUMENT_Brass__SynthBrass_2 = 64 '64   SynthBrass 2
Public Const GM_INSTRUMENT_Brass__Default = GM_INSTRUMENT_Brass__Trumpet

'' Reed     - Reeds
Public Const GM_INSTRUMENT_Reed__Soprano_Sax = 65 '65   Soprano Sax
Public Const GM_INSTRUMENT_Reed__Alto_Sax = 66 '66   Alto Sax
Public Const GM_INSTRUMENT_Reed__Tenor_Sax = 67 '67   Tenor Sax
Public Const GM_INSTRUMENT_Reed__Baritone_Sax = 68 '68   Baritone Sax
Public Const GM_INSTRUMENT_Reed__Oboe = 69 '69   Oboe
Public Const GM_INSTRUMENT_Reed__English_Horn = 70 '70   English Horn
Public Const GM_INSTRUMENT_Reed__Bassoon = 71 '71   Bassoon
Public Const GM_INSTRUMENT_Reed__Clarinet = 72 '72   Clarinet
Public Const GM_INSTRUMENT_Reed__Default = GM_INSTRUMENT_Reed__Soprano_Sax

'' Pipe     - Pipes
Public Const GM_INSTRUMENT_Pipe__Piccolo = 73 '73   Piccolo
Public Const GM_INSTRUMENT_Pipe__Flute = 74 '74   Flute
Public Const GM_INSTRUMENT_Pipe__Recorder = 75 '75   Recorder
Public Const GM_INSTRUMENT_Pipe__Pan_Flute = 76 '76   Pan Flute
Public Const GM_INSTRUMENT_Pipe__Blown_Bottle = 77 '77   Blown Bottle
Public Const GM_INSTRUMENT_Pipe__Skakuhachi = 78 '78   Skakuhachi
Public Const GM_INSTRUMENT_Pipe__Whistle = 79 '79   Whistle
Public Const GM_INSTRUMENT_Pipe__Ocarina = 80 '80   Ocarina
Public Const GM_INSTRUMENT_Pipe__Default = GM_INSTRUMENT_Pipe__Flute

'' Synth Lead     - Synth Leads
Public Const GM_INSTRUMENT_Synth_Lead__Lead_1___square = 81 '81   Lead 1 (square)
Public Const GM_INSTRUMENT_Synth_Lead__Lead_2___sawtooth = 82 '82   Lead 2 (sawtooth)
Public Const GM_INSTRUMENT_Synth_Lead__Lead_3___calliope = 83 '83   Lead 3 (calliope)
Public Const GM_INSTRUMENT_Synth_Lead__Lead_4___chiff = 84 '84   Lead 4 (chiff)
Public Const GM_INSTRUMENT_Synth_Lead__Lead_5___charang = 85 '85   Lead 5 (charang)
Public Const GM_INSTRUMENT_Synth_Lead__Lead_6___voice = 86 '86   Lead 6 (voice)
Public Const GM_INSTRUMENT_Synth_Lead__Lead_7___fifths = 87 '87   Lead 7 (fifths)
Public Const GM_INSTRUMENT_Synth_Lead__Lead_8___bass___lead = 88 '88   Lead 8 (bass+lead)
Public Const GM_INSTRUMENT_Synth_Lead__Default = GM_INSTRUMENT_Synth_Lead__Lead_1___square

'' Synth Pad     - Synth Pads
Public Const GM_INSTRUMENT_Synth_Pad__Pad_1___new_age = 89 '89   Pad 1 (new age)
Public Const GM_INSTRUMENT_Synth_Pad__Pad_2___warm = 90 '90   Pad 2 (warm)
Public Const GM_INSTRUMENT_Synth_Pad__Pad_3___polysynth = 91 '91   Pad 3 (polysynth)
Public Const GM_INSTRUMENT_Synth_Pad__Pad_4___choir = 92 '92   Pad 4 (choir)
Public Const GM_INSTRUMENT_Synth_Pad__Pad_5___bowed = 93 '93   Pad 5 (bowed)
Public Const GM_INSTRUMENT_Synth_Pad__Pad_6___metallic = 94 '94   Pad 6 (metallic)
Public Const GM_INSTRUMENT_Synth_Pad__Pad_7___halo = 95 '95   Pad 7 (halo)
Public Const GM_INSTRUMENT_Synth_Pad__Pad_8___sweep = 96 '96   Pad 8 (sweep)
Public Const GM_INSTRUMENT_Synth_Pad__Default = GM_INSTRUMENT_Synth_Pad__Pad_1___new_age

'' Synth Effects     - Musical Effects
Public Const GM_INSTRUMENT_Synth_Effects__FX_1___rain = 97 '97  FX 1 (rain)
Public Const GM_INSTRUMENT_Synth_Effects__FX_2___soundtrack = 98 '98  FX 2 (soundtrack)
Public Const GM_INSTRUMENT_Synth_Effects__FX_3___crystal = 99 '99  FX 3 (crystal)
Public Const GM_INSTRUMENT_Synth_Effects__FX_4___atmosphere = 100 '100  FX 4 (atmosphere)
Public Const GM_INSTRUMENT_Synth_Effects__FX_5___brightness = 101 '101  FX 5 (brightness)
Public Const GM_INSTRUMENT_Synth_Effects__FX_6___goblins = 102 '102  FX 6 (goblins)
Public Const GM_INSTRUMENT_Synth_Effects__FX_7___echoes = 103 '103  FX 7 (echoes)
Public Const GM_INSTRUMENT_Synth_Effects__FX_8___sci_fi = 104 '104  FX 8 (sci-fi)
Public Const GM_INSTRUMENT_Synth_Effects__Default = GM_INSTRUMENT_Synth_Effects__FX_7___echoes

'' Ethnic     - Ethnic
Public Const GM_INSTRUMENT_Ethnic__Sitar = 105 '105   Sitar
Public Const GM_INSTRUMENT_Ethnic__Banjo = 106 '106   Banjo
Public Const GM_INSTRUMENT_Ethnic__Shamisen = 107 '107   Shamisen
Public Const GM_INSTRUMENT_Ethnic__Koto = 108 '108   Koto
Public Const GM_INSTRUMENT_Ethnic__Kalimba = 109 '109   Kalimba
Public Const GM_INSTRUMENT_Ethnic__Bagpipe = 110 '110   Bagpipe
Public Const GM_INSTRUMENT_Ethnic__Fiddle = 111 '111   Fiddle
Public Const GM_INSTRUMENT_Ethnic__Shanai = 112 '112   Shanai
Public Const GM_INSTRUMENT_Ethnic__Default = GM_INSTRUMENT_Ethnic__Sitar

'' Percussive     - Percussion
Public Const GM_INSTRUMENT_Percussive__Tinkle_Bell = 113 '113   Tinkle Bell
Public Const GM_INSTRUMENT_Percussive__Agogo = 114 '114   Agogo
Public Const GM_INSTRUMENT_Percussive__Steel_Drums = 115 '115   Steel Drums
Public Const GM_INSTRUMENT_Percussive__Woodblock = 116 '116   Woodblock
Public Const GM_INSTRUMENT_Percussive__Taiko_Drum = 117 '117   Taiko Drum
Public Const GM_INSTRUMENT_Percussive__Melodic_Tom = 118 '118   Melodic Tom
Public Const GM_INSTRUMENT_Percussive__Synth_Drum = 119 '119   Synth Drum
Public Const GM_INSTRUMENT_Percussive__Reverse_Cymbal = 120 '120   Reverse Cymbal
Public Const GM_INSTRUMENT_Percussive__Default = GM_INSTRUMENT_Percussive__Tinkle_Bell

''Sound Effects     - Sound Effects
Public Const GM_INSTRUMENT_Sound_Effects__Guitar_Fret_Noise = 121 '121   Guitar Fret Noise
Public Const GM_INSTRUMENT_Sound_Effects__Breath_Noise = 122 '122   Breath Noise
Public Const GM_INSTRUMENT_Sound_Effects__Seashore = 123 '123   Seashore
Public Const GM_INSTRUMENT_Sound_Effects__Bird_Tweet = 124 '124   Bird Tweet
Public Const GM_INSTRUMENT_Sound_Effects__Telephone_Ring = 125 '125   Telephone Ring
Public Const GM_INSTRUMENT_Sound_Effects__Helicopter = 126 '126   Helicopter
Public Const GM_INSTRUMENT_Sound_Effects__Applause = 127 '127   Applause
Public Const GM_INSTRUMENT_Sound_Effects__Gunshot = 128 '128   Gunshot
Public Const GM_INSTRUMENT_Sound_Effects__Default = GM_INSTRUMENT_Sound_Effects__Guitar_Fret_Noise


'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<CONST LIST 2
'***** GM Drum Sounds *****


Public Const GM_DRUM_Sounds__Acoustic_Bass_Drum = 35 '35   Acoustic Bass Drum
Public Const GM_DRUM_Sounds__Bass_Drum_1 = 36 '36   Bass Drum 1
Public Const GM_DRUM_Sounds__Side_Stick = 37 '37    Side Stick
Public Const GM_DRUM_Sounds__Acoustic_Snare = 38 '38    Acoustic Snare
Public Const GM_DRUM_Sounds__Hand_Clap = 39 '39    Hand Clap
Public Const GM_DRUM_Sounds__Electric_Snare = 40 '40    Electric Snare
Public Const GM_DRUM_Sounds__Low_Floor_Tom = 41 '41   Low Floor Tom
Public Const GM_DRUM_Sounds__Closed_Hi_Hat = 42 '42    Closed Hi - Hat
Public Const GM_DRUM_Sounds__High_Floor_Tom = 43 '43   High Floor Tom
Public Const GM_DRUM_Sounds__Pedal_Hi_Hat = 44 '44    Pedal Hi - Hat
Public Const GM_DRUM_Sounds__Low_Tom = 45 '45    Low Tom
Public Const GM_DRUM_Sounds__Open_Hi_Hat = 46 '46   Open Hi-Hat
Public Const GM_DRUM_Sounds__Low_Mid_Tom = 47 '47   Low-Mid Tom
Public Const GM_DRUM_Sounds__Hi_Mid_Tom = 48 '48   Hi-Mid Tom
Public Const GM_DRUM_Sounds__Crash_Cymbal_1 = 49 '49   Crash Cymbal 1
Public Const GM_DRUM_Sounds__High_Tom = 50 '50    High Tom
Public Const GM_DRUM_Sounds__Ride_Cymbal_1 = 51 '51   Ride Cymbal 1
Public Const GM_DRUM_Sounds__Chinese_Cymbal = 52 '52    Chinese Cymbal
Public Const GM_DRUM_Sounds__Ride_Bell = 53 '53    Ride Bell
Public Const GM_DRUM_Sounds__Tambourine = 54 '54    Tambourine
Public Const GM_DRUM_Sounds__Splash_Cymbal = 55 '55    Splash Cymbal
Public Const GM_DRUM_Sounds__Cowbell = 56 '56    Cowbell
Public Const GM_DRUM_Sounds__Crash_Cymbal_2 = 57 '57   Crash Cymbal 2
Public Const GM_DRUM_Sounds__Vibraslap = 58 '58    Vibraslap
Public Const GM_DRUM_Sounds__Ride_Cymbal_2 = 59 '59   Ride Cymbal 2
Public Const GM_DRUM_Sounds__Hi_Bongo = 60 '60   Hi Bongo
Public Const GM_DRUM_Sounds__Low_Bongo = 61 '61   Low Bongo
Public Const GM_DRUM_Sounds__Mute_Hi_Conga = 62 '62   Mute Hi Conga
Public Const GM_DRUM_Sounds__Open_Hi_Conga = 63 '63   Open Hi Conga
Public Const GM_DRUM_Sounds__Low_Conga = 64 '64   Low Conga
Public Const GM_DRUM_Sounds__High_Timbale = 65 '65   High Timbale
Public Const GM_DRUM_Sounds__Low_Timbale = 66 '66   Low Timbale
Public Const GM_DRUM_Sounds__High_Agogo = 67 '67   High Agogo
Public Const GM_DRUM_Sounds__Low_Agogo = 68 '68   Low Agogo
Public Const GM_DRUM_Sounds__Cabasa = 69 '69   Cabasa
Public Const GM_DRUM_Sounds__Maracas = 70 '70   Maracas
Public Const GM_DRUM_Sounds__Short_Whistle = 71 '71   Short Whistle
Public Const GM_DRUM_Sounds__Long_Whistle = 72 '72   Long Whistle
Public Const GM_DRUM_Sounds__Short_Guiro = 73 '73   Short Guiro
Public Const GM_DRUM_Sounds__Long_Guiro = 74 '74   Long Guiro
Public Const GM_DRUM_Sounds__Claves = 75 '75   Claves
Public Const GM_DRUM_Sounds__Hi_Wood_Block = 76 '76   Hi Wood Block
Public Const GM_DRUM_Sounds__Low_Wood_Block = 77 '77   Low Wood Block
Public Const GM_DRUM_Sounds__Mute_Cuica = 78 '78   Mute Cuica
Public Const GM_DRUM_Sounds__Open_Cuica = 79 '79   Open Cuica
Public Const GM_DRUM_Sounds__Mute_Triangle = 80 '80   Mute Triangle
Public Const GM_DRUM_Sounds__Open_Triangle = 81 '81   Open Triangle
Public Const GM_DRUM_Sounds__Default = GM_DRUM_Sounds__Acoustic_Bass_Drum

'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<Routine Area

Public Type Frequence
    Value As Double
End Type

Public Type NoteStr
    Names() As String
    Frequence As Frequence
End Type

Dim lastMIDIError As Long
Dim lastMIDIHandle As Long

Dim notesString As String
Dim notesStr() As String
Dim notesStrCount As Long

Public Function TranslateMIDIDeviceError(ByVal ErrorID As Long) As String
    Select Case ErrorID
        Case API_MMSYSERR_NOERROR
            TranslateMIDIDeviceError = "No Error."
        Case API_MMSYSERR_ERROR
            TranslateMIDIDeviceError = "An Error Occured."
        Case API_MMSYSERR_BADDEVICEID
            TranslateMIDIDeviceError = "Invalid DeviceID."
        Case API_MMSYSERR_NOTENABLED
            TranslateMIDIDeviceError = "Device Not Enabled."
        Case API_MMSYSERR_ALLOCATED
            TranslateMIDIDeviceError = "Device Allocated."
        Case API_MMSYSERR_INVALHANDLE
            TranslateMIDIDeviceError = "Invalid MIDI Handle."
        Case API_MMSYSERR_NODRIVER
            TranslateMIDIDeviceError = "No Driver Found."
        Case API_MMSYSERR_NOMEM
            TranslateMIDIDeviceError = "Not Enough Memory."
        Case API_MMSYSERR_NOTSUPPORTED
            TranslateMIDIDeviceError = "Operation Not Supported."
        Case API_MMSYSERR_BADERRNUM
            TranslateMIDIDeviceError = "Bad Error Number."
        Case API_MMSYSERR_INVALFLAG
            TranslateMIDIDeviceError = "Invalid Initialization Flag(s)."
        Case API_MMSYSERR_INVALPARAM
            TranslateMIDIDeviceError = "Invalid Parameter."
        Case API_MMSYSERR_HANDLEBUSY
            TranslateMIDIDeviceError = "MIDI Device Is Busy."
        Case API_MMSYSERR_INVALIDALIAS
            TranslateMIDIDeviceError = "Invalid Alias."
        Case Else
            TranslateMIDIDeviceError = "An Error Occured."
    End Select
End Function

'-----------------------------------------------------------------------------
'-----------------------------------------------------------------------------
'-----------------------------------------------------------------------------

Public Sub Initialize()
    lastMIDIError = 0
    notesString = "C,C#\Db,D,D#\Eb,E,F,F#\Gb,G,G#\Ab,A,A#\Bb,B"
    notesStr = ProccessNotesStringToNoteArrayString(notesStrCount)
End Sub

Public Function ProccessNotesStringToNoteArrayString(Optional ByRef Count As Long) As String()
    
End Function

'-----------------------------------------------------------------------------
'-----------------------------------------------------------------------------
'-----------------------------------------------------------------------------

Public Sub SendMidiShortOut(hndlMIDI As Long, midiMessage As Long, lowint As Long, highint As Long)
    'Pack MIDI message data into 4 byte long integer
    lowint = (lowint * 256) + midiMessage
    highint = (highint * 256) * 256
    midiMessage = lowint + highint
    'Windows MIDI API function
    Call API_midiOutShortMsg(hndlMIDI, midiMessage)
    lastMIDIHandle = hndlMIDI
End Sub
Public Function NoteToString(ByVal Nr As Long) As String
   Dim octave As Long
   Dim Note As String
   octave = (Nr \ 12)
   Note = Nr Mod 12
   NoteToString = Choose(Note + 1, "C", "C#", "D", "D#", "E", "F", "F#", "G", "G#", "A", "A#", "B") & Format(octave - 1)
End Function
Public Function NoteToStringBass(ByVal Nr As Long) As String
   Dim octave As Long
   Dim Note As String
   octave = (Nr \ 12)
   Note = Nr Mod 12
   NoteToString = Choose(Note + 1, "C", "Db", "D", "Eb", "E", "F", "Gb", "G", "Ab", "A", "Bb", "B") & Format(octave - 1)
End Function

'-----------------------------------------------------------------------------
'-----------------------------------------------------------------------------
'-----------------------------------------------------------------------------
'-----------------------------------MIDI In-----------------------------------
'-----------------------------------------------------------------------------
'-----------------------------------------------------------------------------
'-----------------------------------------------------------------------------

    Public Function MIDIInPortOpen(DeviceID As Long) As Long
        Dim hndlMIDI As Long
        lastMIDIError = API_midiInOpen(hndlMIDI, DeviceID, AddressOf MIDIIn_CallBack, 0, API_CALLBACK_FUNCTION)
        If lastMIDIError = API_MMSYSERR_NOERROR Then
            MIDIInPortOpen = hndlMIDI
        Else
            Call throw("", CLASSID, "MIDIInPortOpen", IIf(lastMIDIError = 0, "Can't Open MIDI In Port.", TranslateMIDIDeviceError(lastMIDIError)))
        End If
        lastMIDIHandle = hndlMIDI
    End Function
    Public Sub MIDIInPortStart(hndlMIDI As Long)
        If hndlMIDI <= 0 Then Call throw("", CLASSID, "MIDIInPortStart", "Invalid MIDI In Handle.")
        lastMIDIError = API_midiInStart(hndlMIDI)
        If lastMIDIError <> API_MMSYSERR_NOERROR Then: Call throw("", CLASSID, "MIDIInPortStart", IIf(lastMIDIError = 0, "Can't Start MIDI In Port.", TranslateMIDIDeviceError(lastMIDIError)))
        lastMIDIHandle = hndlMIDI
    End Sub
    Public Sub MIDIInPortOpenStart(hndlMIDI As Long): Call MIDIInPortOpen(hndlMIDI): Call MIDIInPortStart(hndlMIDI): End Sub
    
    '-----------------------------------------------------------------------------
    '-----------------------------------------------------------------------------
    '-----------------------------------------------------------------------------
    
    Public Sub MIDIInPortStop(hndlMIDI As Long)
        If hndlMIDI <= 0 Then Call throw("", CLASSID, "MIDIInPortStop", "Invalid MIDI In Handle.")
        lastMIDIError = API_midiInStop(hndlMIDI)
        If lastMIDIError <> API_MMSYSERR_NOERROR Then: Call throw("", CLASSID, "MIDIInPortStop", IIf(lastMIDIError = 0, "Can't Stop MIDI In Port.", TranslateMIDIDeviceError(lastMIDIError)))
        lastMIDIHandle = hndlMIDI
    End Sub
    Public Sub MIDIInPortClose(hndlMIDI As Long)
        If hndlMIDI <= 0 Then Call throw("", CLASSID, "MidiInPortClose", "Invalid MIDI In Handle.")
        lastMIDIError = API_midiInClose(hndlMIDI)
        If lastMIDIError <> API_MMSYSERR_NOERROR Then: Call throw("", CLASSID, "MIDIInPortClose", IIf(lastMIDIError = 0, "Can't Close MIDI In Port.", TranslateMIDIDeviceError(lastMIDIError)))
        hndlMIDI = 0
        lastMIDIHandle = hndlMIDI
    End Sub
    Public Sub MIDIInPortStopClose(hndlMIDI As Long): Call MIDIInPortStop(hndlMIDI): Call MIDIInPortClose(hndlMIDI): End Sub

'-----------------------------------------------------------------------------
'-----------------------------------------------------------------------------
'-----------------------------------------------------------------------------
'-----------------------------------MIDI Out----------------------------------
'-----------------------------------------------------------------------------
'-----------------------------------------------------------------------------
'-----------------------------------------------------------------------------

    Public Function MIDIOutPortOpen(DeviceID As Long) As Long
        Dim hndlMIDI As Long
        lastMIDIError = API_midiOutOpen(hndlMIDI, DeviceID, AddressOf MIDIOut_CallBack, 0, API_CALLBACK_FUNCTION)
        If lastMIDIError = API_MMSYSERR_NOERROR Then
            MIDIOutPortOpen = hndlMIDI
        Else
            Call throw("", CLASSID, "MIDIInPortClose", IIf(lastMIDIError = 0, "Can't Open MIDI Out Port.", TranslateMIDIDeviceError(lastMIDIError)))
        End If
        lastMIDIHandle = hndlMIDI
    End Function
    
    '-----------------------------------------------------------------------------
    '-----------------------------------------------------------------------------
    '-----------------------------------------------------------------------------
    
    Public Sub MIDIOutPortClose(hndlMIDI As Long)
        If hndlMIDI <= 0 Then Call throw("", CLASSID, "MidiInPortClose", "Invalid MIDI Out Handle.")
        lastMIDIError = API_midiOutClose(hndlMIDI)
        If lastMIDIError <> API_MMSYSERR_NOERROR Then: Call throw("", CLASSID, "MIDIInPortClose", IIf(lastMIDIError = 0, "Can't Close MIDI In Port.", TranslateMIDIDeviceError(lastMIDIError)))
        hndlMIDI = 0
        lastMIDIHandle = hndlMIDI
    End Sub

'-----------------------------------------------------------------------------
'-----------------------------------------------------------------------------
'-----------------------------------------------------------------------------
'------------------------------------Tools------------------------------------
'-----------------------------------------------------------------------------
'-----------------------------------------------------------------------------
'-----------------------------------------------------------------------------

    Public Sub MIDIConnectInOut(hndlMIDIIn As Long, hndlMIDIOut As Long)
        If (hndlMIDIIn <= 0) Or (hndlMIDIOut <= 0) Then Call throw("", CLASSID, "MIDIConnectInOut", "Invalid MIDI Devices Handle.")
        lastMIDIError = API_midiConnect(hndlMIDIIn, hndlMIDIOut, 0)
        If lastMIDIError <> API_MMSYSERR_NOERROR Then: Call throw("", CLASSID, "MIDIConnectInOut", IIf(lastMIDIError = 0, "Can't Connect MIDI In Device To MIDI Out Device.", TranslateMIDIDeviceError(lastMIDIError)))
    End Sub
    Public Sub MIDIDisconnectInOut(hndlMIDIIn As Long, hndlMIDIOut As Long)
        If (hndlMIDIIn <= 0) Or (hndlMIDIOut <= 0) Then Call throw("", CLASSID, "MIDIDisconnectInOut", "Invalid MIDI Devices Handle.")
        lastMIDIError = API_midiDisconnect(hndlMIDIIn, hndlMIDIOut, 0)
        If lastMIDIError <> API_MMSYSERR_NOERROR Then: Call throw("", CLASSID, "MIDIDisconnectInOut", IIf(lastMIDIError = 0, "Can't Disconnect MIDI In Device From MIDI Out Device.", TranslateMIDIDeviceError(lastMIDIError)))
    End Sub

'-----------------------------------------------------------------------------
'-----------------------------------------------------------------------------
'-----------------------------------------------------------------------------

Public Sub MIDIIn_CallBack(ByVal MIDIInHandle As Long, ByVal wMsg As Long, ByVal dwInstance As Long, ByVal wParam As Long, ByVal lParam As Long)
   Dim txt As String
   Dim Status As Long, OnOff As Long
   Dim NoteNr As Long
   Dim Velocity As Long
   
   On Error Resume Next
   Select Case wMsg
      Case API_MM_MIM_OPEN: txt = "open"
      Case API_MM_MIM_CLOSE: txt = "close"
      Case API_MM_MIM_DATA:
         Status = (wParam Mod 256)
         If Status < &HF0 Then
            Select Case (Status \ 16) ' filter 4-bit channel "n"
               Case &H8, &H9
                  NoteNr = ((wParam \ 256) Mod 256)
                  Velocity = ((wParam \ (256 ^ 2)) Mod 256)
                  OnOff = IIf(Velocity = 0 Or Status = &H80, 0, 1)
                  txt = txt & "Status : Note " & IIf(OnOff = 0, "Off", "On") & vbCrLf
                  txt = txt & "NoteNr : " & NoteToString(NoteNr) & vbCrLf
                  txt = txt & "Velo   :" & str(Velocity)
                  'frmMidi.ShowNote NoteNr, OnOff
               Case &HB
                  txt = "Status : Controller Change"
               Case &HC
                  txt = "Status : Program Change"
               Case &HE
                  txt = "Status : Bender Change"
            End Select
            'Call API_midiOutShortMsg(hMidiOut, wParam)  ' send data = Thru-function
            End If
      Case API_MM_MIM_LONGDATA: txt = "longdata" & " " & Hex(wParam) & " " & Hex(lParam)
      Case API_MM_MIM_ERROR: txt = "error" & " " & Hex(wParam) & " " & Hex(lParam)
      Case API_MM_MIM_LONGERROR: txt = "longerror"
      Case Else: txt = "???"
   End Select
   'If txt <> "" Then frmMidi.lblMidiInfo.Caption = "Midi IN " & vbCrLf & txt
   
End Sub
Public Sub MIDIOut_CallBack(ByVal MIDIOutHandle As Long, ByVal wMsg As Long, ByVal dwInstance As Long, ByVal wParam As Long, ByVal lParam As Long)
    
End Sub
