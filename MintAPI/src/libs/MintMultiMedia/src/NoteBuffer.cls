VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NoteBuffer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Base 0
Option Explicit
Const CLASSID As String = "NoteBuffer"
Const CLASSALIAS As String = APPLICATIONDOMAIN & "." & CLASSID
Const CLASSVERSIONSTRING As String = "0.1"
Const CLASSALIASSTRING As String = CLASSALIAS & "/" & CLASSVERSIONSTRING




Public Enum Instruments
    '***** GM INSTRUMENTS *****

    '' Piano     - Pianos
    INSTRUMENT_Piano__Acoustic_Grand = 1 '1    Acoustic Grand
    INSTRUMENT_Piano__Bright_Acoustic = 2 '2    Bright Acoustic
    INSTRUMENT_Piano__Electric_Grand = 3 '3    Electric Grand
    INSTRUMENT_Piano__Honky_Tonk = 4 '4    Honky -Tonk
    INSTRUMENT_Piano__Electric_Piano_1 = 5 '5    Electric Piano 1
    INSTRUMENT_Piano__Electric_Piano_2 = 6 '6    Electric Piano 2
    INSTRUMENT_Piano__Harpsichord = 7 '7    Harpsichord
    INSTRUMENT_Piano__Clavinet = 8 '8    Clavinet
    INSTRUMENT_Piano__Default = INSTRUMENT_Piano__Acoustic_Grand

    '' Chromatic Percussion     - Tuned Idiophones
    INSTRUMENT_Chromatic_Percussion__Celesta = 9 '9     Celesta
    INSTRUMENT_Chromatic_Percussion__Glockenspiel = 10 '10     Glockenspiel
    INSTRUMENT_Chromatic_Percussion__Music_Box = 11 '11   Music Box
    INSTRUMENT_Chromatic_Percussion__Vibraphone = 12 '12   Vibraphone
    INSTRUMENT_Chromatic_Percussion__Marimba = 13 '13   Marimba
    INSTRUMENT_Chromatic_Percussion__Xylophone = 14 '14   Xylophone
    INSTRUMENT_Chromatic_Percussion__Tubular_Bells = 15 '15   Tubular Bells
    INSTRUMENT_Chromatic_Percussion__Dulcimer = 16 '16   Dulcimer
    INSTRUMENT_Chromatic_Percussion__Default = INSTRUMENT_Chromatic_Percussion__Celesta

    '' Organ     - Organs
    INSTRUMENT_Organ__Drawbar_Organ = 17 '17   Drawbar Organ
    INSTRUMENT_Organ__Percussive_Organ = 18 '18   Percussive Organ
    INSTRUMENT_Organ__Rock_Organ = 19 '19   Rock Organ
    INSTRUMENT_Organ__Church_Organ = 20 '20   Church Organ
    INSTRUMENT_Organ__Reed_Organ = 21 '21   Reed Organ
    INSTRUMENT_Organ__Accoridan = 22 '22   Accoridan
    INSTRUMENT_Organ__Harmonica = 23 '23   Harmonica
    INSTRUMENT_Organ__Tango_Accordian = 24 '24   Tango Accordian
    INSTRUMENT_Organ__Default = INSTRUMENT_Organ__Drawbar_Organ


    '' Guitar     - Guitars
    INSTRUMENT_Guitar__Nylon_String_Guitar = 25 '25   Nylon String Guitar
    INSTRUMENT_Guitar__Steel_String_Guitar = 26 '26   Steel String Guitar
    INSTRUMENT_Guitar__Electric_Jazz_Guitar = 27 '27   Electric Jazz Guitar
    INSTRUMENT_Guitar__Electric_Clean_Guitar = 28 '28   Electric Clean Guitar
    INSTRUMENT_Guitar__Electric_Muted_Guitar = 29 '29   Electric Muted Guitar
    INSTRUMENT_Guitar__Overdriven_Guitar = 30 '30      Overdriven Guitar
    INSTRUMENT_Guitar__Distortion_Guitar = 31 '31      Distortion Guitar
    INSTRUMENT_Guitar__Guitar_Harmonics = 32 '32      Guitar Harmonics
    INSTRUMENT_Guitar__Default = INSTRUMENT_Guitar__Nylon_String_Guitar

    '' Bass     - Basses
    INSTRUMENT_Bass__Acoustic_Bass = 33 '33   Acoustic Bass
    INSTRUMENT_Bass__Electric_Bass___finger = 34 '34   Electric Bass(finger)
    INSTRUMENT_Bass__Electric_Bass___pick = 35 '35   Electric Bass(pick)
    INSTRUMENT_Bass__Fretless_Bass = 36 '36   Fretless Bass
    INSTRUMENT_Bass__Slap_Bass_1 = 37 '37   Slap Bass 1
    INSTRUMENT_Bass__Slap_Bass_2 = 38 '38   Slap Bass 2
    INSTRUMENT_Bass__Synth_Bass_1 = 39 '39   Synth Bass 1
    INSTRUMENT_Bass__Synth_Bass_2 = 40 '40   Synth Bass 2
    INSTRUMENT_Bass__Default = INSTRUMENT_Bass__Electric_Bass___finger

    '' Solo Strings     - String And Timpani
    INSTRUMENT_Solo_Strings__Violin = 41 '41    Violin
    INSTRUMENT_Solo_Strings__Viola = 42 '42    Viola
    INSTRUMENT_Solo_Strings__Cello = 43 '43    Cello
    INSTRUMENT_Solo_Strings__Contrabass = 44 '44    Contrabass
    INSTRUMENT_Solo_Strings__Tremolo_Strings = 45 '45    Tremolo Strings
    INSTRUMENT_Solo_Strings__Pizzicato_Strings = 46 '46    Pizzicato Strings
    INSTRUMENT_Solo_Strings__Orchestral_Strings = 47 '47    Orchestral Strings
    INSTRUMENT_Solo_Strings__Timpani = 48 '48    Timpani
    INSTRUMENT_Solo_Strings__Default = INSTRUMENT_Solo_Strings__Violin

    '' Ensemble Brass     - Ensemble Strings And Voices
    INSTRUMENT_Ensemble__String_Ensemble_1 = 49 '49   String Ensemble 1
    INSTRUMENT_Ensemble__String_Ensemble_2 = 50 '50   String Ensemble 2
    INSTRUMENT_Ensemble__SynthStrings_1 = 51 '51   SynthStrings 1
    INSTRUMENT_Ensemble__SynthStrings_2 = 52 '52   SynthStrings 2
    INSTRUMENT_Ensemble__Choir_Aahs = 53 '53   Choir Aahs
    INSTRUMENT_Ensemble__Voice_Oohs = 54 '54   Voice Oohs
    INSTRUMENT_Ensemble__Synth_Voice = 55 '55   Synth Voice
    INSTRUMENT_Ensemble__Orchestra_Hit = 56 '56   Orchestra Hit
    INSTRUMENT_Ensemble__Default = INSTRUMENT_Ensemble__String_Ensemble_1

    '' Brass     - Brasses
    INSTRUMENT_Brass__Trumpet = 57 '57   Trumpet
    INSTRUMENT_Brass__Trombone = 58 '58   Trombone
    INSTRUMENT_Brass__Tuba = 59 '59   Tuba
    INSTRUMENT_Brass__Muted_Trumpet = 60 '60   Muted Trumpet
    INSTRUMENT_Brass__French_Horn = 61 '61   French Horn
    INSTRUMENT_Brass__Brass_Section = 62 '62   Brass Section
    INSTRUMENT_Brass__SynthBrass_1 = 63 '63    SynthBrass 1
    INSTRUMENT_Brass__SynthBrass_2 = 64 '64   SynthBrass 2
    INSTRUMENT_Brass__Default = INSTRUMENT_Brass__Trumpet

    '' Reed     - Reeds
    INSTRUMENT_Reed__Soprano_Sax = 65 '65   Soprano Sax
    INSTRUMENT_Reed__Alto_Sax = 66 '66   Alto Sax
    INSTRUMENT_Reed__Tenor_Sax = 67 '67   Tenor Sax
    INSTRUMENT_Reed__Baritone_Sax = 68 '68   Baritone Sax
    INSTRUMENT_Reed__Oboe = 69 '69   Oboe
    INSTRUMENT_Reed__English_Horn = 70 '70   English Horn
    INSTRUMENT_Reed__Bassoon = 71 '71   Bassoon
    INSTRUMENT_Reed__Clarinet = 72 '72   Clarinet
    INSTRUMENT_Reed__Default = INSTRUMENT_Reed__Soprano_Sax

    '' Pipe     - Pipes
    INSTRUMENT_Pipe__Piccolo = 73 '73   Piccolo
    INSTRUMENT_Pipe__Flute = 74 '74   Flute
    INSTRUMENT_Pipe__Recorder = 75 '75   Recorder
    INSTRUMENT_Pipe__Pan_Flute = 76 '76   Pan Flute
    INSTRUMENT_Pipe__Blown_Bottle = 77 '77   Blown Bottle
    INSTRUMENT_Pipe__Skakuhachi = 78 '78   Skakuhachi
    INSTRUMENT_Pipe__Whistle = 79 '79   Whistle
    INSTRUMENT_Pipe__Ocarina = 80 '80   Ocarina
    INSTRUMENT_Pipe__Default = INSTRUMENT_Pipe__Flute

    '' Synth Lead     - Synth Leads
    INSTRUMENT_Synth_Lead__Lead_1___square = 81 '81   Lead 1 (square)
    INSTRUMENT_Synth_Lead__Lead_2___sawtooth = 82 '82   Lead 2 (sawtooth)
    INSTRUMENT_Synth_Lead__Lead_3___calliope = 83 '83   Lead 3 (calliope)
    INSTRUMENT_Synth_Lead__Lead_4___chiff = 84 '84   Lead 4 (chiff)
    INSTRUMENT_Synth_Lead__Lead_5___charang = 85 '85   Lead 5 (charang)
    INSTRUMENT_Synth_Lead__Lead_6___voice = 86 '86   Lead 6 (voice)
    INSTRUMENT_Synth_Lead__Lead_7___fifths = 87 '87   Lead 7 (fifths)
    INSTRUMENT_Synth_Lead__Lead_8___bass___lead = 88 '88   Lead 8 (bass+lead)
    INSTRUMENT_Synth_Lead__Default = INSTRUMENT_Synth_Lead__Lead_1___square

    '' Synth Pad     - Synth Pads
    INSTRUMENT_Synth_Pad__Pad_1___new_age = 89 '89   Pad 1 (new age)
    INSTRUMENT_Synth_Pad__Pad_2___warm = 90 '90   Pad 2 (warm)
    INSTRUMENT_Synth_Pad__Pad_3___polysynth = 91 '91   Pad 3 (polysynth)
    INSTRUMENT_Synth_Pad__Pad_4___choir = 92 '92   Pad 4 (choir)
    INSTRUMENT_Synth_Pad__Pad_5___bowed = 93 '93   Pad 5 (bowed)
    INSTRUMENT_Synth_Pad__Pad_6___metallic = 94 '94   Pad 6 (metallic)
    INSTRUMENT_Synth_Pad__Pad_7___halo = 95 '95   Pad 7 (halo)
    INSTRUMENT_Synth_Pad__Pad_8___sweep = 96 '96   Pad 8 (sweep)
    INSTRUMENT_Synth_Pad__Default = INSTRUMENT_Synth_Pad__Pad_1___new_age

    '' Synth Effects     - Musical Effects
    INSTRUMENT_Synth_Effects__FX_1___rain = 97 '97  FX 1 (rain)
    INSTRUMENT_Synth_Effects__FX_2___soundtrack = 98 '98  FX 2 (soundtrack)
    INSTRUMENT_Synth_Effects__FX_3___crystal = 99 '99  FX 3 (crystal)
    INSTRUMENT_Synth_Effects__FX_4___atmosphere = 100 '100  FX 4 (atmosphere)
    INSTRUMENT_Synth_Effects__FX_5___brightness = 101 '101  FX 5 (brightness)
    INSTRUMENT_Synth_Effects__FX_6___goblins = 102 '102  FX 6 (goblins)
    INSTRUMENT_Synth_Effects__FX_7___echoes = 103 '103  FX 7 (echoes)
    INSTRUMENT_Synth_Effects__FX_8___sci_fi = 104 '104  FX 8 (sci-fi)
    INSTRUMENT_Synth_Effects__Default = INSTRUMENT_Synth_Effects__FX_7___echoes

    '' Ethnic     - Ethnic
    INSTRUMENT_Ethnic__Sitar = 105 '105   Sitar
    INSTRUMENT_Ethnic__Banjo = 106 '106   Banjo
    INSTRUMENT_Ethnic__Shamisen = 107 '107   Shamisen
    INSTRUMENT_Ethnic__Koto = 108 '108   Koto
    INSTRUMENT_Ethnic__Kalimba = 109 '109   Kalimba
    INSTRUMENT_Ethnic__Bagpipe = 110 '110   Bagpipe
    INSTRUMENT_Ethnic__Fiddle = 111 '111   Fiddle
    INSTRUMENT_Ethnic__Shanai = 112 '112   Shanai
    INSTRUMENT_Ethnic__Default = INSTRUMENT_Ethnic__Sitar

    '' Percussive     - Percussion
    INSTRUMENT_Percussive__Tinkle_Bell = 113 '113   Tinkle Bell
    INSTRUMENT_Percussive__Agogo = 114 '114   Agogo
    INSTRUMENT_Percussive__Steel_Drums = 115 '115   Steel Drums
    INSTRUMENT_Percussive__Woodblock = 116 '116   Woodblock
    INSTRUMENT_Percussive__Taiko_Drum = 117 '117   Taiko Drum
    INSTRUMENT_Percussive__Melodic_Tom = 118 '118   Melodic Tom
    INSTRUMENT_Percussive__Synth_Drum = 119 '119   Synth Drum
    INSTRUMENT_Percussive__Reverse_Cymbal = 120 '120   Reverse Cymbal
    INSTRUMENT_Percussive__Default = INSTRUMENT_Percussive__Tinkle_Bell
    ''Sound Effects     - Sound Effects
    INSTRUMENT_Sound_Effects__Guitar_Fret_Noise = 121 '121   Guitar Fret Noise
    INSTRUMENT_Sound_Effects__Breath_Noise = 122 '122   Breath Noise
    INSTRUMENT_Sound_Effects__Seashore = 123 '123   Seashore
    INSTRUMENT_Sound_Effects__Bird_Tweet = 124 '124   Bird Tweet
    INSTRUMENT_Sound_Effects__Telephone_Ring = 125 '125   Telephone Ring
    INSTRUMENT_Sound_Effects__Helicopter = 126 '126   Helicopter
    INSTRUMENT_Sound_Effects__Applause = 127 '127   Applause
    INSTRUMENT_Sound_Effects__Gunshot = 128 '128   Gunshot
    INSTRUMENT_Sound_Effects__Default = INSTRUMENT_Sound_Effects__Guitar_Fret_Noise
    ''Others     - DrumsN
    INSTRUMENT_Others__Drums = 100 '135   Drum (Must Plus With GM Drum Sounds.)
End Enum
Public Enum Drums
    '***** GM Drum Sounds *****
    DRUM_Sounds__Acoustic_Bass_Drum = 35 '35   Acoustic Bass Drum
    DRUM_Sounds__Bass_Drum_1 = 36 '36   Bass Drum 1
    DRUM_Sounds__Side_Stick = 37 '37    Side Stick
    DRUM_Sounds__Acoustic_Snare = 38 '38    Acoustic Snare
    DRUM_Sounds__Hand_Clap = 39 '39    Hand Clap
    DRUM_Sounds__Electric_Snare = 40 '40    Electric Snare
    DRUM_Sounds__Low_Floor_Tom = 41 '41   Low Floor Tom
    DRUM_Sounds__Closed_Hi_Hat = 42 '42    Closed Hi - Hat
    DRUM_Sounds__High_Floor_Tom = 43 '43   High Floor Tom
    DRUM_Sounds__Pedal_Hi_Hat = 44 '44    Pedal Hi - Hat
    DRUM_Sounds__Low_Tom = 45 '45    Low Tom
    DRUM_Sounds__Open_Hi_Hat = 46 '46   Open Hi-Hat
    DRUM_Sounds__Low_Mid_Tom = 47 '47   Low-Mid Tom
    DRUM_Sounds__Hi_Mid_Tom = 48 '48   Hi-Mid Tom
    DRUM_Sounds__Crash_Cymbal_1 = 49 '49   Crash Cymbal 1
    DRUM_Sounds__High_Tom = 50 '50    High Tom
    DRUM_Sounds__Ride_Cymbal_1 = 51 '51   Ride Cymbal 1
    DRUM_Sounds__Chinese_Cymbal = 52 '52    Chinese Cymbal
    DRUM_Sounds__Ride_Bell = 53 '53    Ride Bell
    DRUM_Sounds__Tambourine = 54 '54    Tambourine
    DRUM_Sounds__Splash_Cymbal = 55 '55    Splash Cymbal
    DRUM_Sounds__Cowbell = 56 '56    Cowbell
    DRUM_Sounds__Crash_Cymbal_2 = 57 '57   Crash Cymbal 2
    DRUM_Sounds__Vibraslap = 58 '58    Vibraslap
    DRUM_Sounds__Ride_Cymbal_2 = 59 '59   Ride Cymbal 2
    DRUM_Sounds__Hi_Bongo = 60 '60   Hi Bongo
    DRUM_Sounds__Low_Bongo = 61 '61   Low Bongo
    DRUM_Sounds__Mute_Hi_Conga = 62 '62   Mute Hi Conga
    DRUM_Sounds__Open_Hi_Conga = 63 '63   Open Hi Conga
    DRUM_Sounds__Low_Conga = 64 '64   Low Conga
    DRUM_Sounds__High_Timbale = 65 '65   High Timbale
    DRUM_Sounds__Low_Timbale = 66 '66   Low Timbale
    DRUM_Sounds__High_Agogo = 67 '67   High Agogo
    DRUM_Sounds__Low_Agogo = 68 '68   Low Agogo
    DRUM_Sounds__Cabasa = 69 '69   Cabasa
    DRUM_Sounds__Maracas = 70 '70   Maracas
    DRUM_Sounds__Short_Whistle = 71 '71   Short Whistle
    DRUM_Sounds__Long_Whistle = 72 '72   Long Whistle
    DRUM_Sounds__Short_Guiro = 73 '73   Short Guiro
    DRUM_Sounds__Long_Guiro = 74 '74   Long Guiro
    DRUM_Sounds__Claves = 75 '75   Claves
    DRUM_Sounds__Hi_Wood_Block = 76 '76   Hi Wood Block
    DRUM_Sounds__Low_Wood_Block = 77 '77   Low Wood Block
    DRUM_Sounds__Mute_Cuica = 78 '78   Mute Cuica
    DRUM_Sounds__Open_Cuica = 79 '79   Open Cuica
    DRUM_Sounds__Mute_Triangle = 80 '80   Mute Triangle
    DRUM_Sounds__Open_Triangle = 81 '81   Open Triangle
    DRUM_Sounds__Default = DRUM_Sounds__Acoustic_Bass_Drum
End Enum


Public Enum NoteBufferBufferPlaces
    nbMemory = 0
    nbSourceStorageLocation = 1

    nbHardDisk = 5
End Enum
Public Enum NoteTypes
    ntMIDI
    ntRoNOTE
End Enum

Friend Sub Initialize(Optional Path As String = "")
    
End Sub

Public Function LoadFromFile(Path As String) As NoteBuffer
    
End Function
Public Function LoadFromMIDIInput(ByVal MIDIInput As MIDIInput, Optional NumberOfNotesToRead, Optional ByVal Timeout As Long = mcInfinite) As NoteBuffer
    
End Function
Public Function Clone() As NoteBuffer
    
End Function

Public Sub Load(bufferplace As NoteBufferBufferPlaces)
    
End Sub

Public Property Get Instrument(TrackIndex As Long) As Instruments
    
End Property
Public Property Get Caption(TrackIndex As Long) As String
    
End Property
Public Property Get Volume(TrackIndex As Long) As Long
    
End Property

Public Property Get EnumTracks() As Long
    
End Property
Public Property Get EnumBars() As Long
    
End Property

Public Property Get State() As Long
    
End Property

Public Function IsLoaded() As Boolean
    
End Function
Public Function IsLoadedF() As Boolean
    
End Function

Public Function CanPlay() As Boolean
    
End Function
Public Function Length() As Long
    
End Function


Private Property Get IAliasable_Alias() As String
    IAliasable_Alias = CLASSALIASSTRING
End Property

Private Function IClassClonable_Clone() As Object
    Set IClassClonable_Clone = Clone
End Function

