VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "libMethods"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@PROJECT_LICENSE
Option Explicit
Option Base 0
Const CLASSID As String = "baseMethods"

Private Type API_baseMethods_OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128      '  Maintenance string for PSS usage
End Type

Private Declare Function API_baseMethods_VarPtrArray Lib "msvbvm60" Alias "VarPtr" (Ptr() As Any) As Long
Private Declare Sub API_baseMethods_CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function API_baseMethods_GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As API_baseMethods_OSVERSIONINFO) As Long
Private Declare Function API_baseMethods_GetVersion Lib "kernel32" Alias "GetVersion" () As Long
Private Declare Function API_baseMethods_GetUserName Lib "advapi32" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function API_baseMethods_GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function API_baseMethods_GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function API_baseMethods_GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function API_baseMethods_GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function API_baseMethods_SetComputerName Lib "kernel32" Alias "SetComputerNameA" (ByVal lpComputerName As String) As Long
Private Declare Function API_baseMethods_GetUserDefaultLangID Lib "kernel32" Alias "GetUserDefaultLangID" () As Integer
Private Declare Function API_baseMethods_GetSystemDefaultLCID Lib "kernel32" Alias "GetSystemDefaultLCID" () As Long
Private Declare Function API_baseMethods_GetUserDefaultLCID Lib "kernel32" Alias "GetUserDefaultLCID" () As Long
Private Declare Function API_baseMethods_SetCurrentDirectory Lib "kernel32" Alias "SetCurrentDirectoryA" (ByVal lpPathName As String) As Long
Private Declare Function API_baseMethods_GetCurrentDirectory Lib "kernel32" Alias "GetCurrentDirectory" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
'Private Declare Function API_baseMethods_CreateDirectory Lib "kernel32" Alias "CreateDirectoryA" (ByVal lpPathName As String, lpSecurityAttributes As SECURITY_ATTRIBUTES) As Long
'Private Declare Function API_baseMethods_CreateDirectoryEx Lib "kernel32" Alias "CreateDirectoryExA" (ByVal lpTemplateDirectory As String, ByVal lpNewDirectory As String, lpSecurityAttributes As SECURITY_ATTRIBUTES) As Long
Private Declare Function API_baseMethods_RemoveDirectory Lib "kernel32" Alias "RemoveDirectoryA" (ByVal lpPathName As String) As Long
Private Declare Function API_baseMethods_GetFullPathName Lib "kernel32" Alias "GetFullPathNameA" (ByVal lpFileName As String, ByVal nBufferLength As Long, ByVal lpBuffer As String, ByVal lpFilePart As String) As Long
Private Declare Function API_baseMethods_GetUserProfileDirectory Lib "userenv" (ByVal hToken As Long, ByVal lpProfileDir As String, ByRef lpcchSize As Long) As Long
Private Declare Function API_baseMethods_GetSiteDirectory Lib "advapi32" Alias "GetSiteDirectoryA" (ByVal hToken As Long, ByVal pszSiteDirectory As String, ByVal uSize As Long) As Long
Private Declare Function API_baseMethods_GetSystemWindowsDirectory Lib "kernel32" Alias "GetSystemWindowsDirectoryA" (ByVal lpBuffer As String, ByVal uSize As Long) As Long
Private Declare Function API_baseMethods_GetProfilesDirectory Lib "userenv" Alias "GetProfilesDirectoryA" (ByVal lpProfileDir As String, lpcchSize As Long) As Boolean

Public Enum ConcatPathEndMode
    cpeEndsWithSlash
    cpeNoEndsWithSlash
    cpeNotMatter
End Enum

Public Enum equalToFlags
    AllMustEqual = 1
    SomeEqual = 2
    SomeNotEqual = 4
    AllMustNotEqual = 8
    
    SomeCompare = SomeEqual Or SomeNotEqual
    AllCompare = AllMustEqual Or AllMustNotEqual
    NotValue = 32
    AllValues = AllMustEqual Or SomeEqual Or SomeNotEqual Or AllMustNotEqual
End Enum

Public Function ArraySize(targetArray) As Long
If Not (VarType(targetArray) And vbArray) = vbArray Then throw InvalidArgumentTypeException
    On Error GoTo zeroLength
    ArraySize = (UBound(targetArray) - LBound(targetArray) + 1)
zeroLength:
End Function
Public Function IsArrayEmpty(targetArray) As Boolean
If Not (VarType(targetArray) And vbArray) = vbArray Then throw InvalidArgumentTypeException
    On Error GoTo zeroLength
    IsArrayEmpty = (UBound(targetArray) - LBound(targetArray) + 1) <= 0
    Exit Function
zeroLength:
    IsArrayEmpty = True
End Function
Public Sub EmptyArray(targetArray)
If Not (VarType(targetArray) And vbArray = vbArray) Then throw InvalidArgumentTypeException
    Erase targetArray
End Sub
Public Function IsEmptyVariable(targetVariable) As Boolean
    Select Case VarType(targetVariable)
        Case vbArray
            IsEmptyVariable = (IsArrayEmpty(targetVariable))
        Case vbObject
            IsEmptyVariable = (targetVariable Is Nothing)
        Case Else
            IsEmptyVariable = (targetVariable = Empty)
    End Select
End Function

Public Function new_(arg)
    If VarType(arg) = vbObject Then
        Set new_ = arg
    Else
            new_ = arg
    End If
End Function

Public Function GetLPSTR(LPSTR As String) As String
    Dim Index As Long
    Index = InStr(1, LPSTR, Chr(0)) - 1
    If Index <= 0 Then Exit Function
    GetLPSTR = Left(LPSTR, Index)
End Function
Public Function GetLPSTRWL(LPSTR As String, Length As Long) As String 'With Length
    GetLPSTRWL = Left(LPSTR, Length)
End Function

'Check If EnumerateValue Has One Of The flags Conditions When Compairing It To args.
'ex: if equalTo(SomeEqual,x,10,29,4) then //means if x equal to one of the 10,29,4 returns true else false
'flags:Compaire Mode.
'EnumerateValue:Value To Compaire.
'args:Array To Compaire With EnumerateValue.
'returns:bool
Public Function equalTo(flags As equalToFlags, EnumerateValue, ParamArray args() As Variant) As Boolean
    Dim i As Long
    On Error GoTo errZeroArgs
    i = LBound(args)
    GoTo notZeroArgs
errZeroArgs:
    If (flags And AllMustNotEqual = AllMustNotEqual) Or (flags And SomeNotEqual = SomeNotEqual) Then
        equalTo = (EnumerateValue <> 0)
    Else
        equalTo = (EnumerateValue = 0)
    End If
    Exit Function
notZeroArgs:
    On Error Resume Next
    Select Case VarType(EnumerateValue)
    Case VbVarType.vbArray 'VarType(EnumerateValue)
        Select Case flags And equalToFlags.AllValues
            Case equalToFlags.AllMustEqual: For i = LBound(args) To UBound(args)
                If Not ArrayCompare(EnumerateValue, args(i)) Then: equalTo = False: Exit Function
            Next: equalTo = True
            Case equalToFlags.AllMustNotEqual: For i = LBound(args) To UBound(args)
                If ArrayCompare(EnumerateValue, args(i)) Then: equalTo = False: Exit Function
            Next: equalTo = True
            Case equalToFlags.SomeEqual: For i = LBound(args) To UBound(args)
                If ArrayCompare(EnumerateValue, args(i)) Then: equalTo = True: Exit Function
            Next: equalTo = False
            Case equalToFlags.SomeNotEqual: For i = LBound(args) To UBound(args)
                If Not ArrayCompare(EnumerateValue, args(i)) Then: equalTo = True: Exit Function
            Next: equalTo = False
            Case Else: throw UnknownValueException("Unknown Flags argument value.")
        End Select 'Flags And equalToFlags.AllValues
    Case VbVarType.vbObject 'VarType(EnumerateValue)
        Select Case flags And equalToFlags.AllValues
            Case equalToFlags.AllMustEqual: For i = LBound(args) To UBound(args)
                If Not EnumerateValue Is args(i) Then: equalTo = False: Exit Function
            Next: equalTo = True
            Case equalToFlags.AllMustNotEqual: For i = LBound(args) To UBound(args)
                If EnumerateValue Is args(i) Then: equalTo = False: Exit Function
            Next: equalTo = True
            Case equalToFlags.SomeEqual: For i = LBound(args) To UBound(args)
                If EnumerateValue Is args(i) Then: equalTo = True: Exit Function
            Next: equalTo = False
            Case equalToFlags.SomeNotEqual: For i = LBound(args) To UBound(args)
                If Not EnumerateValue Is args(i) Then: equalTo = True: Exit Function
            Next: equalTo = False
            Case Else: throw UnknownValueException("Unknown Flags argument value.")
        End Select 'Flags And equalToFlags.AllValues
    Case Else 'VarType(EnumerateValue)
        Select Case flags And equalToFlags.AllValues
            Case equalToFlags.AllMustEqual: For i = LBound(args) To UBound(args)
                If EnumerateValue <> CLng(args(i)) Then: equalTo = False: Exit Function
            Next: equalTo = True
            Case equalToFlags.AllMustNotEqual: For i = LBound(args) To UBound(args)
                If EnumerateValue = CLng(args(i)) Then: equalTo = False: Exit Function
            Next: equalTo = True
            Case equalToFlags.SomeEqual: For i = LBound(args) To UBound(args)
                If EnumerateValue = CLng(args(i)) Then: equalTo = True: Exit Function
            Next: equalTo = False
            Case equalToFlags.SomeNotEqual: For i = LBound(args) To UBound(args)
                If EnumerateValue <> CLng(args(i)) Then: equalTo = True: Exit Function
            Next: equalTo = False
            Case Else: throw UnknownValueException("Unknown Flags argument value.")
        End Select 'Flags And equalToFlags.AllValues
    End Select 'VarType(EnumerateValue)
    If flags And NotValue = NotValue Then equalTo = Not equalTo
End Function
'-------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------------
Public Function GetVersion() As Long
    GetVersion = API_baseMethods_GetVersion
End Function
Public Function CurrentUser() As String
    Dim buf As String * SMALLLPSTR, bufSize As Long
    bufSize = SMALLLPSTR
    buf = String(SMALLLPSTR, Chr(0))
    Call API_baseMethods_GetUserName(buf, bufSize)
    CurrentUser = GetLPSTR(buf)
End Function
Public Function GetTempPath() As String
    Dim buf As String * LARGELPSTR, bufSize As Long
    bufSize = LARGELPSTR
    buf = String(LARGELPSTR, Chr(0))
    Call API_baseMethods_GetTempPath(bufSize, buf)
    GetTempPath = GetLPSTR(buf)
End Function
Public Function GetSystemPath() As String
    Dim buf As String * LARGELPSTR, bufSize As Long
    bufSize = LARGELPSTR
    buf = String(LARGELPSTR, Chr(0))
    Call API_baseMethods_GetSystemDirectory(buf, bufSize)
    GetSystemPath = GetLPSTR(buf)
End Function
Public Function GetWindowsPath() As String
    Dim buf As String * LARGELPSTR, bufSize As Long
    bufSize = LARGELPSTR
    buf = String(LARGELPSTR, Chr(0))
    Call API_baseMethods_GetWindowsDirectory(buf, bufSize)
    GetWindowsPath = GetLPSTR(buf)
End Function
Public Function GetUserPath() As String
    Dim buf As String * SMALLLPSTR, bufSize As Long
    bufSize = LARGELPSTR
    buf = String(SMALLLPSTR, Chr(0))
    Call API_baseMethods_GetProfilesDirectory(buf, bufSize)
    GetUserPath = GetLPSTR(buf)
End Function
'-------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------------
Public Function EnumDirectoryFiles(Path As String, Optional Attributes As VbFileAttribute) As String()
    Dim cp As String, Count As Long, retVal() As String
    If Not (Right(Path, 1) = "/" Or Right(Path, 1) = "\") Then Path = Path & "\"
    cp = Dir(ConcatPath(Path, "", cpeEndsWithSlash), vbNormal Or Attributes)
    While cp <> ""
        ReDim Preserve retVal(Count)
        retVal(Count) = cp
        Count = Count + 1
        cp = Dir
    Wend
    EnumDirectoryFiles = retVal
End Function
Public Function CountDirectoryFiles(Path As String, Optional Attributes As VbFileAttribute) As Long
    Dim cp As String, Count As Long
    If Not (Right(Path, 1) = "/" Or Right(Path, 1) = "\") Then Path = Path & "\"
    cp = Dir(ConcatPath(Path, "", cpeEndsWithSlash), Attributes)
    While cp <> ""
        Count = Count + 1
        cp = Dir
    Wend
    CountDirectoryFiles = Count
End Function
Public Function EnumDirectoryFolders(Path As String, Optional Attributes As VbFileAttribute = 0) As String()
    Dim cp As String, Count As Long, retVal() As String
    If Not (Right(Path, 1) = "/" Or Right(Path, 1) = "\") Then Path = Path & "\"
    cp = Dir(Path, vbDirectory Or Attributes)
    If cp <> "" Then cp = Dir
    If cp <> "" Then cp = Dir
    While cp <> ""
        On Local Error GoTo cantOpenFile
        If (GetAttr(ConcatPath(Path, cp)) And vbDirectory) = vbDirectory Then
            On Error GoTo 0
            ReDim Preserve retVal(Count)
            retVal(Count) = cp
            Count = Count + 1
        End If
cantOpenFile:
        cp = Dir
    Wend
    EnumDirectoryFolders = retVal
End Function
Public Function CountDirectoryFolders(Path As String, Optional Attributes As VbFileAttribute = 0) As Long
    Dim cp As String, Count As Long
    If Not (Right(Path, 1) = "/" Or Right(Path, 1) = "\") Then Path = Path & "\"
    cp = Dir(Path, vbDirectory Or Attributes)
    If cp <> "" Then cp = Dir
    If cp <> "" Then cp = Dir
    While cp <> ""
        On Local Error GoTo cantOpenFile
        If (GetAttr(ConcatPath(Path, cp)) And vbDirectory) = vbDirectory Then
            On Error GoTo 0
            Count = Count + 1
        End If
cantOpenFile:
        cp = Dir
    Wend
    CountDirectoryFolders = Count
End Function
Public Function CountSubDirectoriesFiles(Path As String, Optional Attributes As VbFileAttribute = 0) As Long
    Dim Count As Long, subDirs() As String, subCount As Long
    If Not (Right(Path, 1) = "/" Or Right(Path, 1) = "\") Then Path = Path & "\"
    subDirs = EnumDirectoryFolders(Path, Attributes)
    subCount = ArraySize(subDirs)
    Dim i As Long
    Count = CountDirectoryFiles(Path, Attributes)
    For i = 0 To subCount - 1
        Count = Count + CountSubDirectoriesFiles(ConcatPath(Path, subDirs(i), cpeEndsWithSlash), Attributes)
    Next
    CountSubDirectoriesFiles = Count
End Function
Public Function EnumSubDirectoriesFiles(Path As String, Optional Attributes As VbFileAttribute = 0) As String()
    Dim subDirs() As String, subCount As Long, retVal() As String, bufVal() As String
    If Not (Right(Path, 1) = "/" Or Right(Path, 1) = "\") Then Path = Path & "\"
    subDirs = EnumDirectoryFolders(Path, Attributes)
    subCount = ArraySize(subDirs)
    Dim i As Long
    retVal = EnumDirectoryFiles(Path, Attributes)
    For i = 0 To subCount - 1
        bufVal = EnumSubDirectoriesFiles(ConcatPath(Path, subDirs(i), cpeEndsWithSlash), Attributes)
        Call AppendArrayToArray(retVal, bufVal)
    Next
    EnumSubDirectoriesFiles = retVal()
End Function
Public Function GetFilePath(ByVal Path As String) As String
On Error GoTo err
    Dim slashIndex As Long, backslashIndex As Long
    slashIndex = InStrRev(Path, "/")
    backslashIndex = InStrRev(Path, "\")
    slashIndex = IIf(slashIndex >= backslashIndex, slashIndex, backslashIndex)
    If slashIndex = 0 Then throw InvalidPathException
    backslashIndex = Len(Path)
    GetFilePath = Left(Path, backslashIndex - (backslashIndex - slashIndex))
Exit Function
err:
End Function
Public Function GetFileName(ByVal Path As String) As String
On Error GoTo err
    Dim slashIndex As Long, backslashIndex As Long
    slashIndex = InStrRev(Path, "/")
    backslashIndex = InStrRev(Path, "\")
    slashIndex = IIf(slashIndex >= backslashIndex, slashIndex, backslashIndex)
    If slashIndex = 0 Then throw InvalidPathException
    backslashIndex = Len(Path)
    GetFileName = Right(Path, backslashIndex - slashIndex)
Exit Function
err:
End Function
Public Function GetFileExtention(ByVal Path As String) As String
On Error GoTo err
    Dim slashIndex As Long, backslashIndex As Long
    slashIndex = InStrRev(Path, "/")
    backslashIndex = InStrRev(Path, "\")
    slashIndex = IIf(slashIndex >= backslashIndex, slashIndex, backslashIndex)
    'If slashIndex = 0 Then throw InvalidPathException
    backslashIndex = InStrRev(Path, ".")
    If backslashIndex = 0 Then
        GetFileExtention = ""
        Exit Function
    End If
    If slashIndex > backslashIndex Then
        GetFileExtention = ""
        Exit Function
    Else
        slashIndex = Len(Path)
        GetFileExtention = Right(Path, slashIndex - backslashIndex)
    End If
Exit Function
err:
End Function
Public Function GetFileNameOnly(ByVal Path As String) As String
On Error GoTo err
    Dim slashIndex As Long, backslashIndex As Long, fLen As Long
    fLen = Len(Path)
    slashIndex = InStrRev(Path, "/")
    backslashIndex = InStrRev(Path, "\")
    slashIndex = IIf(slashIndex >= backslashIndex, slashIndex, backslashIndex)
    If slashIndex = 0 Then throw InvalidPathException
    backslashIndex = InStrRev(Path, ".")
    If backslashIndex = 0 Then
        GetFileNameOnly = Right(Path, fLen - slashIndex)
        Exit Function
    End If
    If slashIndex > backslashIndex Then
        GetFileNameOnly = Right(Path, fLen - slashIndex)
        Exit Function
    Else
        GetFileNameOnly = Mid(Path, slashIndex + 1, backslashIndex - slashIndex - 1)
    End If
Exit Function
err:
End Function
Public Function ConcatPath(ByVal Path, ByVal PathToAdd, Optional EndWithSlash As ConcatPathEndMode = ConcatPathEndMode.cpeNotMatter, Optional Slash As String = "\") As String
    Dim p As String, A As String
    A = Trim(PathToAdd)
    p = Trim(Path)
    If A <> "" Then
        If Left(A, 1) = Slash Then
            A = Mid(A, 2)
        End If
    Else
        ConcatPath = p
        GoTo checkLastSlash
    End If
    If p <> "" Then
        If Right(p, 1) <> Slash Then
            p = p & Slash
        End If
    Else
        ConcatPath = A
        GoTo checkLastSlash
    End If
    ConcatPath = p & A
checkLastSlash:
    If Not equalTo(SomeEqual, EndWithSlash, ConcatPathEndMode.cpeNoEndsWithSlash, ConcatPathEndMode.cpeEndsWithSlash) Then Exit Function
    If ConcatPath <> "" Then
        If Right(ConcatPath, 1) = Slash Then
            If EndWithSlash = cpeNoEndsWithSlash Then ConcatPath = Mid(ConcatPath, 1, Len(ConcatPath) - 1)
        Else
            If EndWithSlash = cpeEndsWithSlash Then ConcatPath = ConcatPath & Slash
        End If
    End If
End Function
Public Function CheckPathValidation(Path As String, MakeTrueForm As Boolean, Optional CheckForDrive As Boolean = True) As Boolean
    If MakeTrueForm Then
        Path = Replace(Path, "/", "\")
    End If
    CheckPathValidation = True
End Function
'-------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------------
'MemoryToByteArray : copies specified memory address value to a bytearray.
'targetAddress : the address of memory to copy to byte array.
'SourceSize : source memory address content size to copy to byte array.
'Times : determines the times that source memory value with length equals to SourceSize copies to byte array.
Public Function MemoryToByteArray(ByVal targetAddress As Long, SourceSize As Long, Optional times As Long = 1) As Byte()
    Dim outLen As Long
    outLen = (SourceSize * times) - 1
    If outLen <= 0 Then throw InvalidArgumentValueException
    Dim retVal() As Byte
    ReDim retVal(outLen)
    
    Dim c_byte_value As Byte
    
    Dim i As Long
    For i = 0 To outLen - 1
        Call API_baseMethods_CopyMemory(ByVal targetAddress + i, ByVal VarPtr(c_byte_value), 1)
        retVal(i) = c_byte_value
    Next

    MemoryToByteArray = retVal
End Function
'ByteArrayToMemory : copies specified byte array to memory target's address.
'targetArray : source array to copy to memory.
'BytesToCopy : number of bytes used to copy byte array.
Public Sub ByteArrayToMemory(targetArray, BytesToCopy As Long)
    
End Sub
Public Function FillArray(targetArray, whattofill, OutputSize As Long, Optional FillToLeft = True)
    If (VarType(targetArray) And vbArray) <> vbArray Then throw InvalidArgumentTypeException
    Dim targetLength As Long
    targetLength = ArraySize(targetArray)
    
End Function
Public Function StringToByteArray(str As String, Optional Length As Long = -1) As Byte()
    If Length = 0 Then Exit Function
    Dim strLen As Long
    Dim retVal() As Byte
    strLen = Len(str)
    If strLen = 0 Then Exit Function
    If Length < 0 Then Length = strLen
    ReDim retVal(Length - 1)
    Dim i As Long
    For i = 1 To Length
        retVal(i - 1) = Asc(Mid(str, i, 1))
    Next
    StringToByteArray = retVal()
End Function
Public Function ByteArrayToString(b() As Byte, Optional Length As Long = -1) As String
    If Length = 0 Then Exit Function
    Dim str As String, arrsize As Long
    On Error GoTo zeroLength
    arrsize = UBound(b) - LBound(b) + 1
zeroLength:
    If arrsize = 0 Then Exit Function
    If Length < 0 Then Length = arrsize
    Dim i As Long
    For i = LBound(b) To UBound(b)
        str = str & Chr(b(i))
    Next
    ByteArrayToString = str
End Function
Public Function LongToByteArray(lngNum As Long) As Byte()
    Dim b(3) As Byte
    b(0) = lngNum & &HFF000000
    b(1) = lngNum & &HFF0000
    b(2) = lngNum & &HFF00
    b(3) = lngNum & &HFF
    LongToByteArray = b
End Function
Public Function ByteArrayToLong(b() As Byte) As Long
    Dim size As Long, retVal As Long, buf As Long
    size = ArraySize(b)
    If size <> 4 Then
        If size <= 0 Or size > 4 Then
            throw OutOfRangeException
        Else
            Call FillArray(b, 0, 4, False)
        End If
    End If
    retVal = ShiftLeft(b(0), 24)
    buf = ShiftLeft(b(1), 16): retVal = retVal + buf
    buf = ShiftLeft(b(2), 8): retVal = retVal + buf
    retVal = retVal + b(3)
    ByteArrayToLong = retVal
End Function
Public Function IntegerToByteArray(intNum As Integer) As Byte()
    Dim b(1) As Byte
    b(0) = intNum & &HFF00
    b(1) = intNum & &HFF
    IntegerToByteArray = b
End Function
Public Function ByteArrayToInteger(b() As Byte) As Integer
    Dim size As Long, retVal As Integer, buf As Integer
    size = ArraySize(b)
    If size <= 0 Or size > 2 Then throw OutOfRangeException
    retVal = ShiftLeft(b(0), 24)
    If size > 1 Then buf = ShiftLeft(b(1), 16): retVal = retVal + buf
    ByteArrayToInteger = retVal
End Function
Public Function DateToByteArray(intNum As Date) As Byte()
    throw NotImplementedException
End Function
Public Function ByteArrayToDate(b() As Byte) As Date
    Dim size As Long, retVal As Long, buf As Long
    size = ArraySize(b)
    If size <> 8 Then
        If size <= 0 Or size > 8 Then
            throw OutOfRangeException
        Else
            Call FillArray(b, 0, 8, False)
        End If
    End If
    retVal = ShiftLeft(b(0), 56)
    buf = ShiftLeft(b(1), 48): retVal = retVal + buf
    buf = ShiftLeft(b(2), 40): retVal = retVal + buf
    buf = ShiftLeft(b(3), 32): retVal = retVal + buf
    buf = ShiftLeft(b(4), 24): retVal = retVal + buf
    buf = ShiftLeft(b(5), 16): retVal = retVal + buf
    buf = ShiftLeft(b(6), 8): retVal = retVal + buf
    retVal = retVal + b(7)
    ByteArrayToDate = retVal
End Function
Public Function DoubleToByteArray(intNum As Date) As Byte()
    throw NotImplementedException
End Function
Public Function ByteArrayToDouble(b() As Byte) As Date
    Dim size As Long, retVal As Long, buf As Long
    size = ArraySize(b)
    If size <> 8 Then
        If size <= 0 Or size > 8 Then
            throw OutOfRangeException
        Else
            Call FillArray(b, 0, 8, False)
        End If
    End If
    retVal = ShiftLeft(b(0), 56)
    buf = ShiftLeft(b(1), 48): retVal = retVal + buf
    buf = ShiftLeft(b(2), 40): retVal = retVal + buf
    buf = ShiftLeft(b(3), 32): retVal = retVal + buf
    buf = ShiftLeft(b(4), 24): retVal = retVal + buf
    buf = ShiftLeft(b(5), 16): retVal = retVal + buf
    buf = ShiftLeft(b(6), 8): retVal = retVal + buf
    retVal = retVal + b(7)
    ByteArrayToDouble = retVal
End Function

Public Sub ByteArrayToArgs(b() As Byte, args() As Variant)
    
End Sub
Public Function ArgsToByteArray(args() As Variant) As Byte()
    
End Function

Public Sub InsertArrayIndex(targetArray, Index, Item)
    If Not (VarType(targetArray) And vbArray) Then throw InvalidArgumentTypeException("Only Arrays Accepted.")
    
End Sub
Public Sub AppendToArray(sourceArray, Item)
    Dim sourceUBound As Long
    If (VarType(sourceArray) And vbArray) <> vbArray Then throw InvalidArgumentTypeException
    If Not IsArrayEmpty(sourceArray) Then
        sourceUBound = UBound(sourceArray)
        sourceUBound = sourceUBound + 1
    Else
        sourceUBound = 0
    End If
    ReDim Preserve sourceArray(sourceUBound)
    If VarType(Item) = vbObject Then
        Set sourceArray(sourceUBound) = Item
    Else
            sourceArray(sourceUBound) = Item
    End If
End Sub
Public Sub AppendArrayToArray(sourceArray, targetArray)
    If (VarType(sourceArray) And vbArray) <> vbArray Then throw InvalidArgumentTypeException
    If (VarType(sourceArray) And vbArray) <> vbArray Then
        Call AppendToArray(sourceArray, targetArray)
        Exit Sub
    End If
    Dim sourceUBound As Long, sourceLBound As Long
    Dim targetUBound As Long, targetLBound As Long
    If IsArrayEmpty(targetArray) Then
        Exit Sub
    Else
        If IsArrayEmpty(sourceArray) Then
            sourceArray = targetArray
        End If
    End If
    sourceLBound = LBound(sourceArray)
    sourceUBound = UBound(sourceArray)
    targetLBound = LBound(targetArray)
    targetUBound = UBound(targetArray)
    Dim i As Long
    For i = targetLBound To targetUBound
        sourceUBound = sourceUBound + 1
        ReDim Preserve sourceArray(sourceUBound)
        If VarType(targetArray(i)) = vbObject Then
            Set sourceArray(sourceUBound) = targetArray(i)
        Else
                sourceArray(sourceUBound) = targetArray(i)
        End If
    Next
End Sub

Public Function ExpandVArrayToString(targetArray, Optional throwExceptionOnUnknownValues As Boolean = True) As String
    Dim targetArrayType As VbVarType, i As Long, retVal As String
    targetArrayType = VarType(targetArray)
    For i = LBound(targetArray) To UBound(targetArray)
        If (targetArrayType And vbArray) = vbArray Then
            retVal = retVal & ExpandVArrayToString(targetArray(i), throwExceptionOnUnknownValues)
        ElseIf (targetArrayType And vbObject) = vbObject Then
            If throwExceptionOnUnknownValues Then throw InvalidArgumentTypeException("One Or More Array Items Type Is Object.")
        Else
            retVal = retVal & CStr(targetArray(i))
        End If
    Next
    ExpandVArrayToString = retVal
End Function
'-------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------------
Public Function BinaryCompare(a1, a2, Optional LengthToCompare As Long = -1) As Boolean
    Dim a1Type As VbVarType
    a1Type = VarType(a1)
    If a1Type <> VarType(a2) Then: Exit Function
    If a1Type = vbArray Then
        BinaryCompare = ArrayCompare(a1, a2, LengthToCompare)
    ElseIf a1Type = vbObject Then
        BinaryCompare = (a1 Is a2)
    Else
        BinaryCompare = (a1 = a2)
    End If
End Function
Public Function ArrayCompare(a1, a2, Optional LengthToCompare As Long = -1) As Boolean
    Dim a1Type As VbVarType, a2Type As VbVarType
    a1Type = VarType(a1): a2Type = VarType(a2)
    If Not (((a1Type And vbArray) = vbArray) And ((a2Type And vbArray) = vbArray)) Then throw InvalidArgumentTypeException("Only Arrays Accepted.")
    If (a1Type <> a2Type) Then throw InvalidArgumentTypeException("a1 Array Type Must Equal To a2 Array Type.")
    Dim a1Len As Long, a2Len As Long
    On Error GoTo a1_zeroLength
    a1Len = UBound(a1) - LBound(a1) + 1
a1_zeroLength:
    On Error GoTo a2_zeroLength
    a2Len = UBound(a2) - LBound(a2) + 1
a2_zeroLength:
    If (LBound(a2) <> LBound(a1)) Then GoTo some_returnFalse
    Dim ln As Long
    If LengthToCompare = -1 Then
        If a1Len <> a2Len Then GoTo some_returnFalse
        LengthToCompare = a1Len
    End If
    If a1Len < LengthToCompare Then GoTo some_returnFalse
    If a2Len < LengthToCompare Then GoTo some_returnFalse
    ln = LengthToCompare - 1
    On Error GoTo 0
    Dim i As Long
    For i = LBound(a1) To ln
        If a1(i) <> a2(i) Then GoTo some_returnFalse
    Next
    ArrayCompare = True
    Exit Function
some_returnFalse:
    ArrayCompare = False
'    Exit Function
'notZeroStart:
'    throw InvalidArrayLBoundException
End Function

Public Function RepeatString(str As String, Count As Long, Optional Splitter As String = "") As String
    Dim i As Long
    For i = 1 To Count
        RepeatString = RepeatString & str & Splitter
    Next
End Function

