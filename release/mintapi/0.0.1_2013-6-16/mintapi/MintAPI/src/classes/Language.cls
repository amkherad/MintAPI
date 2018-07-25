VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Language"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Base 0
Option Explicit
Const CLASSID As String = "Language"
Const CLASSALIAS As String = APPLICATIONDOMAIN & "." & CLASSID
Const CLASSVERSIONSTRING As String = "0.1"
Const CLASSALIASSTRING As String = CLASSALIAS & "/" & CLASSVERSIONSTRING

'*.lang,*.lng,*.mlang,*.mlng,*.minttranslation

Implements IAliasable


Const LANGUAGEFILEVALIDATIONP1 As Long = 1953392973 ' Mint
Const LANGUAGEFILEVALIDATIONP2 As Long = 1735287116 ' Lang
Const LANGFILEVERSION As Long = &H100001 '&H100001

Private Type trRecord
    KeyLength As Long
    KeyValue As String
    TranslateLength As Long
    TranslateValue As String
End Type

Dim i_name As String '50B
Dim i_key As String 'shortname '5B
Dim i_region As String '20B
Dim i_description As String 'dynamicLength
Dim i_righttoleft As Boolean

Dim rec() As trRecord
Dim recCount As Long

Dim lngPath As String
Dim i_isloaded As Boolean
Dim i_isvirtualloaded As Boolean

Private Sub Class_Initialize()
    i_name = "English"
    i_key = "en"
    i_region = "United States"
    i_description = "English (United States) | Global English Language."
    i_righttoleft = False
End Sub

Friend Sub Initialize(Optional Path As String = "")
    lngPath = Path
End Sub

Friend Sub SetAttributes(Name As String, Key As String, Region As String, Description As String, RightToLeft As Boolean)
    i_name = Name
    i_key = Key
    i_region = Region
    i_description = Description
    i_righttoleft = RightToLeft
End Sub
Friend Function CreateLanguage(Name As String, Key As String, Region As String, Description As String, RightToLeft As Boolean) As Language
    Dim l As New Language
    Call l.SetAttributes(Name, Key, Region, Description, RightToLeft)
    Set CreateLanguage = l
End Function

Public Sub Unload()
    Call Dispose
End Sub
Public Function LogicalLoad() As Language
    If i_isloaded Then throw Exception("Language Already Loaded.")
    i_isloaded = True
    i_isvirtualloaded = True
    Set LogicalLoad = Me
End Function
Public Function Load() As Language
    If i_isloaded Then throw Exception("Language Already Loaded.")
    Call LoadFromFile_internal(lngPath)
    Set Load = Me
End Function
Public Function LoadFromFile(Optional OptionalPath) As Language
    Dim VL As String
    VL = IIf(IsMissing(OptionalPath), lngPath, CStr(OptionalPath))
    Dim l As New Language
    Call l.LoadFromFile_internal(VL)
    Set LoadFromFile = l
End Function
Friend Sub LoadFromFile_internal(Path As String)
On Error GoTo Err
    Call Dispose

    Dim fl As Long, VL As Long ', f As File
    Dim rName As String * 50, rKey As String * 5, rRegion As String * 20

    Dim intArr() As Integer
    Dim i As Long
    
    'Set f = File(Path)
    'Call f.OpenFile(fmOpen, fNormal, faRead, fshRead)
    fl = FreeFile
    Open Path For Binary As #fl
    
    'f.ReadData VL
    Get #fl, , VL
    If VL <> LANGUAGEFILEVALIDATIONP1 Then throw InvalidFileException
    Get #fl, , VL
    If VL <> LANGUAGEFILEVALIDATIONP2 Then throw InvalidFileException
    Get #fl, , VL
    If VL <> LANGFILEVERSION Then throw InvalidFileException("Unknown language file version.")

    Get #fl, , rName
    Get #fl, , rKey
    Get #fl, , rRegion
    Get #fl, , i_righttoleft

    Dim i_description_length As Long
    Get #fl, , i_description_length
    If i_description_length > 0 Then
        ReDim intArr(i_description_length - 1)
        Get #fl, , intArr
        i_description = IntegerArrayToString(intArr)
    End If

    i_name = Trim(rName)
    i_key = Trim(rKey)
    i_region = Trim(rRegion)

    Get #fl, , recCount 'number of records
    ReDim rec(recCount - 1)
    For i = 0 To recCount - 1
        Get #fl, , VL
        If VL <= 0 Then GoTo Err
        ReDim intArr(VL - 1)
        Get #fl, , intArr
        If ArraySize(intArr) <> VL Then GoTo Err

        rec(i).KeyLength = VL
        rec(i).KeyValue = IntegerArrayToString(intArr)

        Get #fl, , VL
        If VL > 0 Then
            ReDim intArr(VL - 1)
            Get #fl, , intArr
            If ArraySize(intArr) <> VL Then GoTo Err

            rec(i).TranslateLength = VL
            rec(i).TranslateValue = IntegerArrayToString(intArr)
        End If
    Next
    Close #fl
    i_isloaded = True
    Exit Sub
Err:
    Close #fl
    Call Dispose
    throw InvalidFileException("Invalid Language File.")
End Sub
Friend Sub LoadFromFile_internal1(Path As String)
On Error GoTo Err
    Call Dispose

    Dim f As File
    Dim VL As Long
    Dim rName As String * 50, rKey As String * 5, rRegion As String * 20

    Dim intArr() As Integer
    Dim i As Long

    Set f = File(Path).OpenFile(fomRead)

    Call f.ReadData(VL)
    If VL <> LANGUAGEFILEVALIDATIONP1 Then throw InvalidFileException
    Call f.ReadData(VL)
    If VL <> LANGUAGEFILEVALIDATIONP2 Then throw InvalidFileException
    Call f.ReadData(VL)
    If VL <> LANGFILEVERSION Then throw InvalidFileException("Unknown language file version.")
    
    Call f.ReadData(rName)
    Call f.ReadData(rKey)
    Call f.ReadData(rRegion)
    Call f.ReadData(i_righttoleft)

    Dim i_description_length As Long
    Call f.ReadData(i_description_length)
    If i_description_length > 0 Then
        ReDim intArr(i_description_length - 1)
        Call f.ReadData(intArr)
        i_description = IntegerArrayToString(intArr)
    End If

    i_name = Trim(rName)
    i_key = Trim(rKey)
    i_region = Trim(rRegion)
    Call f.ReadData(recCount) 'number of records
    ReDim rec(recCount - 1)
    For i = 0 To recCount - 1
        Call f.ReadData(VL)
        If VL <= 0 Then GoTo Err
        ReDim intArr(VL - 1)
        Call f.ReadData(intArr)
        If ArraySize(intArr) <> VL Then GoTo Err

        rec(i).KeyLength = VL
        rec(i).KeyValue = IntegerArrayToString(intArr)

        Call f.ReadData(VL)
        If VL > 0 Then
            ReDim intArr(VL - 1)
            Call f.ReadData(intArr)
            If ArraySize(intArr) <> VL Then GoTo Err

            rec(i).TranslateLength = VL
            rec(i).TranslateValue = IntegerArrayToString(intArr)
        End If
    Next
    Call f.CloseFile
    i_isloaded = True
    Exit Sub
Err:
    Call f.CloseFile
    Call Dispose
    throw InvalidFileException("Invalid Language File.")
End Sub
Public Sub CompileTranslationFile(Path As String)
    Dim fl As Long, VL As Long
    Dim rName As String * 50, rKey As String * 5, rRegion As String * 20

    rName = i_name
    rKey = i_key
    rRegion = i_region

    Dim intArr() As Integer
    Dim i As Long

    If FileExists(Path) Then Call Kill(Path)

    fl = FreeFile
    Open Path For Binary As #fl

    Put #fl, , LANGUAGEFILEVALIDATIONP1
    Put #fl, , LANGUAGEFILEVALIDATIONP2
    Put #fl, , LANGFILEVERSION

    Put #fl, , rName
    Put #fl, , rKey
    Put #fl, , rRegion
    Put #fl, , i_righttoleft

    Dim i_description_length As Long
    i_description_length = Len(i_description)

    Put #fl, , i_description_length
    If i_description_length > 0 Then
        intArr = StringToIntegerArray(i_description)
        Put #fl, , intArr
    End If

    Put #fl, , recCount 'number of records
    For i = 0 To recCount - 1
        Put #fl, , rec(i).KeyLength
        intArr = StringToIntegerArray(rec(i).KeyValue)
        Put #fl, , intArr

        Put #fl, , rec(i).TranslateLength
        If rec(i).TranslateLength > 0 Then
            intArr = StringToIntegerArray(rec(i).TranslateValue)
            Put #fl, , intArr
        End If
    Next
    Close #fl
    Exit Sub
Err:
    Close #fl
    Call Dispose
    throw InvalidFileException
End Sub

Friend Sub Dispose()
    i_name = ""
    i_key = ""
    i_region = ""
    i_description = ""
    i_righttoleft = False
    i_isloaded = False
    i_isvirtualloaded = False

    Call removeAll
End Sub

Public Property Get Name() As String
    Name = i_name
End Property
Public Property Get ShortName() As String
    ShortName = i_key
End Property
Public Property Get Region() As String
    Region = i_region
End Property
Public Property Get Description() As String
    Description = i_description
End Property
Public Property Get RightToLeft() As Boolean
    RightToLeft = i_righttoleft
End Property

Public Property Get IsLoaded() As Boolean
    IsLoaded = i_isloaded
End Property
Public Property Get Path() As String
    Path = lngPath
End Property


Friend Function IndexOf(Key As String) As Long
    Dim i As Long
    For i = 0 To recCount - 1
        If rec(i).KeyValue = Key Then
            IndexOf = i
            Exit Function
        End If
    Next
    IndexOf = -1
End Function
Friend Function translationExists(Key As String) As Boolean
    translationExists = (IndexOf(Key) >= 0)
End Function
Friend Sub appendTranslate(Key As String, translate As String)
    If Key = "" Then throw InvalidArgumentValueException
    If IndexOf(Key) >= 0 Then throw ItemExistsException
    ReDim Preserve rec(recCount)
    rec(recCount).KeyLength = Len(Key)
    rec(recCount).KeyValue = Key
    rec(recCount).TranslateLength = Len(translate)
    rec(recCount).TranslateValue = translate
    recCount = recCount + 1
End Sub
Friend Function translationKey(Index As Long) As String
    If Index < 0 Or Index >= recCount Then throw OutOfRangeException
    translationKey = rec(Index).KeyValue
End Function
Friend Function translationValue(Index As Long) As String
    If Index < 0 Or Index >= recCount Then throw OutOfRangeException
    translationValue = rec(Index).TranslateValue
End Function
Friend Property Get Count() As Long
    Count = recCount
End Property
Friend Function removeTranslate(Key As String) As Boolean
    If recCount = 0 Then GoTo exitWithNoMercy
    Dim Index As Long
    Index = IndexOf(Key)
    If Index = -1 Then GoTo exitWithNoMercy
    Dim i As Long
    recCount = recCount - 1
    For i = Index To recCount - 1
        rec(i) = rec(i + 1)
    Next
    If recCount > 0 Then
        ReDim Preserve rec(recCount - 1)
    Else
        Erase rec
    End If
    removeTranslate = True
    Exit Function
exitWithNoMercy:
    removeTranslate = False
End Function
Friend Sub removeAll()
    recCount = 0
    Erase rec
End Sub
Friend Sub changeTranslate(Key As String, newTranslate As String)
    Dim Index As Long
    Index = IndexOf(Key)
    If Index = -1 Then throw ItemNotExistsException
    rec(Index).TranslateValue = newTranslate
    rec(Index).TranslateLength = Len(newTranslate)
End Sub

Public Function Exists(Key As String) As Boolean
    Exists = (IndexOf(Key) >= 0)
End Function
Public Function tr(Key As String) As String
    Dim Index As Long
    Index = IndexOf(Key)
    If Index = -1 Then
        tr = Key
    Else
        If rec(Index).TranslateValue = "" Then
            tr = Key
        Else
            tr = rec(Index).TranslateValue
        End If
    End If
End Function
Public Function translate(Key As String) As String
    Dim Index As Long
    Index = IndexOf(Key)
    If Index = -1 Then
        translate = Key
    Else
        If rec(Index).TranslateValue = "" Then
            translate = Key
        Else
            translate = rec(Index).TranslateValue
        End If
    End If
End Function


Friend Sub CopyFrom(Language As Language)
    Dim i As Long
    Call removeAll
    i_name = Language.Name
    i_region = Language.Region
    i_key = Language.ShortName
    i_righttoleft = Language.RightToLeft
    i_description = Language.Description
    i_isloaded = Language.IsLoaded
    lngPath = Language.Path
    For i = 0 To Language.Count - 1
        Call appendTranslate(Language.translationKey(i), Language.translationValue(i))
    Next
End Sub
Friend Sub CopyRecordsFrom(Language As Language)
    Dim i As Long
    Call removeAll
    For i = 0 To Language.Count - 1
        Call appendTranslate(Language.translationKey(i), Language.translationValue(i))
    Next
End Sub

Friend Sub appendTo(Language As Language)
    Dim i As Long
    For i = 0 To recCount - 1
        Call Language.appendTranslate(rec(i).KeyValue, rec(i).TranslateValue)
    Next
End Sub

Private Property Get IAliasable_Alias() As String
    IAliasable_Alias = CLASSALIASSTRING
End Property
