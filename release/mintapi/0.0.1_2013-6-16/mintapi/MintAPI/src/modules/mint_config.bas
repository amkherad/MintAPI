Attribute VB_Name = "mint_config"
Option Explicit

Private Const CONFIG_VALIDATIONKEY01  As Long = 1953392973 'Mint
Private Const CONFIG_VALIDATIONKEY02  As Long = 1734766147 'Cnfg

Private Const VIRTUAL_CONFIG_PATH As String = "config.bin"

Private Const MIN_VARNAME As Long = 20
Private Const MIN_VARVALUE As Long = 75
Private Const MIN_ARRAY_ITEMS As Long = 5
Private Const MAX_CHECKSUM As Long = 32
Private Const MAX_BLOCKNAME As Long = 11
Private Const MAX_LOGFILEPATH As Long = 256

Private Const VTBL_CONFIG_VALIDATION01 As Long = 1
Private Const VTBL_CONFIG_VALIDATION02 As Long = VTBL_CONFIG_VALIDATION01 + 4
Private Const VTBL_CONFIG_VERSION As Long = VTBL_CONFIG_VALIDATION02 + 4
Private Const VTBL_CONFIG_CHECKSUM As Long = VTBL_CONFIG_VERSION + 4 '256 bit '32 byte
Private Const VTBL_CONFIG_BLOCKS As Long = VTBL_CONFIG_CHECKSUM + MAX_CHECKSUM
'Private Const VTBL_CONFIG_LOGFILEPATH As Long = VTBL_CONFIG_CHECKSUM + 32

Private Const MINT_VNG As String = "_mint"

Private Const CNFG_VARNAME_LICENSESCOUNT As String = MINT_VNG & "_lic_count"
Private Const CNFG_VARNAME_LICENSES As String = MINT_VNG & "_lic"

Private Const CNFG_VARNAME_TEMPPATH As String = MINT_VNG & "_temppath"
Private Const CNFG_VARNAME_LOGPATH As String = MINT_VNG & "_logpath"
Private Const CNFG_VARNAME_PLUGINSPATH As String = MINT_VNG & "_pluginspath"
Private Const CNFG_VARNAME_DATAPATH As String = MINT_VNG & "_datapath"
Private Const CNFG_VARNAME_CONFIGPATH As String = MINT_VNG & "_configpath"

Public Type CONFIG_LPSTRING
    Length As Long
    strValue As String
End Type

Dim is_config_streaming As Boolean

Dim streaming_file As File

'{FILETYPE}{FILETYPE2}
'{Mint Version}{Last Update}

Public Sub Initialize()
    If Not FileExists(ConfigPath) Then _
        Call CreateEmptyConfigFile(ConfigPath)


End Sub
Public Sub Dispose(Optional ByVal Force As Boolean = False)

End Sub


Public Sub DllLoadConfiguration()
    Dim i As Long, retStr As String
    '=========================
    Dim LicenseCount As Long
    LicenseCount = ReadMintAPIDecimalVariable(CNFG_VARNAME_LICENSESCOUNT)
    For i = 1 To LicenseCount
        On Error Resume Next
        Dim LicenseStr As String, LInfoStr As String, VLineIndex As Long
        retStr = ReadMintAPIStringVariable(CNFG_VARNAME_LICENSES & "[" & i & "]")
        VLineIndex = InStr(1, retStr, "|")
        If VLineIndex > 0 Then
            LicenseStr = Left(retStr, VLineIndex - 1)
            LInfoStr = mID(retStr, VLineIndex + 1)
            If Not licensing.RegisterLibraryLicense(LicenseStr, LInfoStr) Then
                Call RemoveArrayedVariables(CNFG_VARNAME_LICENSES, LicenseCount, i)
            End If
        Else
            Call RemoveArrayedVariables(CNFG_VARNAME_LICENSES, LicenseCount, i)
        End If
    Next
    On Error GoTo 0
    '=========================

End Sub

'Private Function OpenConfigFile() As File
'    If is_config_streaming Then
'        Set OpenConfigFile = streaming_file
'        Exit Function
'    End If
'    Dim OpenConfigFile_retVal As New File
'    Set OpenConfigFile_retVal = File(ConfigPath)
'    Call OpenConfigFile_retVal.OpenFile(fmOpen, fNormal, faWrite, fshReadWrite)
'    Set OpenConfigFile = OpenConfigFile_retVal
'End Function
Public Sub BeginConfigEditing()
    If streaming_file Is Nothing Then   'throw InvalidStatusException("BeginConfigEditing() already called.")
        Dim OpenConfigFile_retVal As New File
        Set OpenConfigFile_retVal = File(ConfigPath)
        Call OpenConfigFile_retVal.OpenFile(fmOpen, fNormal, faRead, fshReadWrite)
        Set streaming_file = OpenConfigFile_retVal
    End If
    is_config_streaming = True
End Sub
Public Sub EndConfigEditing()
    If Not streaming_file Is Nothing Then  'throw InvalidStatusException("BeginConfigEditing() must be called before.")
        Call streaming_file.CloseFile 'it flushes data automaticly.
        Set streaming_file = Nothing '...
    End If
    is_config_streaming = False
End Sub

Public Function ValidateMintAPIVariableName(Name As String) As Boolean
    If Left(Name, Len(MINT_VNG)) = MINT_VNG Then Exit Function
    If InStr(Name, Chr(0)) > 0 Then Exit Function
    Dim ascChr As Byte
    ascChr = Asc(Left(Name, 1))

    If Not (Left(Name, 1) = "_") Then
        If Not ((ascChr >= 65) And (ascChr <= 90) Or _
            (ascChr >= 97) And (ascChr <= 122)) Then Exit Function
    End If

    'If InStr(Name, " ") > 0 Then Exit Function
    ValidateMintAPIVariableName = True
End Function

Public Sub RemoveArrayedVariables(AName As String, CountAll As Long, RemoveIndex As Long)

End Sub

Public Function ReadMintAPIDecimalVariable(Name As String) As Long

End Function
Public Function WriteMintAPIDecimalVariable(Name As String) As Long

End Function
Public Function ReadMintAPIDoubleVariable(Name As String) As Long

End Function
Public Function WriteMintAPIDoubleVariable(Name As String) As Long

End Function
Public Function ReadMintAPIStringVariable(Name As String) As String

End Function
Public Function WriteMintAPIStringVariable(Name As String) As String

End Function

Public Function CountMintAPIVariables() As Long

End Function
Public Function CheckMintAPIVariable(Name As String, ByRef Output() As Byte) As Boolean

End Function
Public Function ReadMintAPIVariable(Name As String, Optional NotFoundError As Boolean = True) As Byte()

End Function
Public Function WriteMintAPIVariable(Name As String, Value() As Byte)

End Function
Public Sub DeleteMintAPIVariable(Name As String)

End Sub

'=================================================================
' MintAPI Configuration Methods.
'=================================================================

Private Sub ReadLPSTRING(File As File, lpString As CONFIG_LPSTRING, Optional Index As Long = -1)
    If Not File.IsOpen Then throw InvalidStatusException("File not opened.")
    If lpString.Length <= 0 Then Exit Sub
    Dim bt() As Byte
    ReDim bt(lpString.Length - 1)
    If Index = -1 Then
        Call File.ReadData(bt)
    Else
        Dim posBuffer As Long
        posBuffer = File.Position
        File.Position = Index
        Call File.ReadData(bt)
        File.Position = posBuffer
    End If
    lpString.strValue = ByteArrayToString(bt)
End Sub
Private Sub WriteLPSTRING(File As File, lpString As CONFIG_LPSTRING, Optional Index As Long = -1)
    If Not File.IsOpen Then throw InvalidStatusException("File not opened.")
    If lpString.Length <= 0 Then Exit Sub
    Dim bt() As Byte
    bt = StringToByteArray(lpString.strValue)
    If Index = -1 Then
        Call File.WriteData(lpString.Length)
        Call File.WriteData(bt)
    Else
        Dim posBuffer As Long
        posBuffer = File.Position
        File.Position = Index
        Call File.WriteData(lpString.Length)
        Call File.WriteData(bt)
        File.Position = posBuffer
    End If
End Sub

Private Function GenerateCheckSum() As Byte()
    Dim retVal() As Byte
    ReDim retVal(MAX_CHECKSUM - 1)
    GenerateCheckSum = retVal
End Function

Private Sub CreateEmptyConfigFile(Path As String)
    Dim f As File
    Set f = File(Path)
On Error GoTo ErrHandler
    Call f.OpenFile(fmCreate, fNormal, faReadWrite, fshRead)
    f.Position = VTBL_CONFIG_VALIDATION01
    Call f.WriteData(CONFIG_VALIDATIONKEY01)

    f.Position = VTBL_CONFIG_VALIDATION02
    Call f.WriteData(CONFIG_VALIDATIONKEY02)

    f.Position = VTBL_CONFIG_VERSION
    Call f.WriteData(MintAPIVersion)

    f.Position = VTBL_CONFIG_CHECKSUM
    Call f.WriteData(GenerateCheckSum)
ErrHandler:
On Error GoTo Err
    Call f.CloseFile
Err:
End Sub











Public Function ConfigPath() As String
    ConfigPath = ConcatPath(App.Path, VIRTUAL_CONFIG_PATH)
End Function
Public Function MintAPIDataPath() As String
    BeginConfigEditing
On Error GoTo ErrorHandler
    Dim baArr() As Byte
    If Not CheckMintAPIVariable(CNFG_VARNAME_DATAPATH, baArr) Then
        MintAPIDataPath = ConcatPath(App.Path, "data")
    Else
        MintAPIDataPath = Trim$(ByteArray(baArr).toLPSTR)
    End If
ErrorHandler:
    EndConfigEditing
End Function
Public Function MintAPILogPath() As String
    BeginConfigEditing
On Error GoTo ErrorHandler
    Dim baArr() As Byte
    If Not CheckMintAPIVariable(CNFG_VARNAME_LOGPATH, baArr) Then
        MintAPILogPath = ConcatPath(MintAPITempPath, "log.txt")
    Else
        MintAPILogPath = Trim$(ByteArray(baArr).toLPSTR)
    End If
ErrorHandler:
    EndConfigEditing
End Function
Public Function MintAPIPluginsPath() As String
    BeginConfigEditing
On Error GoTo ErrorHandler
    Dim baArr() As Byte
    If Not CheckMintAPIVariable(CNFG_VARNAME_PLUGINSPATH, baArr) Then
        MintAPIPluginsPath = ConcatPath(App.Path, "plugins")
    Else
        MintAPIPluginsPath = Trim$(ByteArray(baArr).toLPSTR)
    End If
ErrorHandler:
    EndConfigEditing
End Function
Public Function MintAPITempPath() As String
    BeginConfigEditing
On Error GoTo ErrorHandler
    Dim baArr() As Byte
    If Not CheckMintAPIVariable(CNFG_VARNAME_TEMPPATH, baArr) Then
        MintAPITempPath = GetApplicationSpecifiedTempPath_specified _
                         (App.CompanyName, APPLICATIONID, App.Major, _
                          App.Minor, App.Revision, False)
    Else
        MintAPITempPath = Trim$(ByteArray(baArr).toLPSTR)
    End If
ErrorHandler:
    EndConfigEditing
End Function

