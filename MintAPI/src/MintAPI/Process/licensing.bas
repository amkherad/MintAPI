Attribute VB_Name = "licensing"
Option Explicit

Const MINCOUNT As Long = 15
Const MAXCOUNT As Long = 30

Public Enum ProductType
    ptDesign = &H1
    ptConfiguration = &H2
    ptCore = &H4
    ptGraphics = &H8
    ptPlugins = &H10
    ptAdvStringParser = &H20
    ptAdvCore = &H40
    ptAdvGraphics = &H80
    ptNoteBuffer = &H100 '***** editing notes <> NOT BUFFERING NOTES *****
    ptMIDIDevice = &H200


    pt_C_Personal = ptConfiguration Or ptDesign Or ptGraphics
    pt_C_Application = ptConfiguration Or ptAdvStringParser Or ptDesign Or ptCore Or ptGraphics Or ptMIDIDevice
    pt_C_Commercial = pt_C_Application Or ptPlugins Or ptAdvGraphics Or ptMIDIDevice
    pt_C_Enterprise = pt_C_Application Or ptAdvGraphics Or ptAdvCore Or ptNoteBuffer Or ptMIDIDevice
    pt_C_Ultimate = pt_C_Enterprise Or pt_C_Commercial Or pt_C_Application Or pt_C_Personal
End Enum

Private Type License
    keyCount As Long
    Key() As Byte
    Info As String
    License As String
    PType As ProductType
End Type
Dim Licenses() As License
Dim LicensesCount As Long


Public Sub Construct()
    Call RegisterLibraryLicense("0000000000000000000000000", "MintAPI Learning License Version.")
    
End Sub
Public Sub Dispose(Optional ByVal Force As Boolean = False)
    
End Sub

'Private Function getch(b As Byte) As String
'    Dim u As String * 1
'    b = b - MINCOUNT
'    If b >= 26 Then b = b + 6
'    If b > 66 Then throw Exps.InvalidArgumentValueException
'    u = Chr(65 + (b))
'    getch = u
'End Function
'Private Function getasc(ByVal B As Byte) As Byte
'    B = B - MINCOUNT
'    If B >= 26 Then B = B + 6
'    If B > 66 Then throw Exps.InvalidArgumentValueException
'    getasc = 65 + (B)
'End Function
'Private Function getchr(ByVal B As Long) As Long
'    If B > 66 Then throw Exps.InvalidArgumentValueException
'    B = B - 65
'    B = B + 15
'    If B >= 26 Then B = B - 6
'    If B < 0 Then throw Exps.InvalidArgumentValueException
'    getchr = B
'End Function
'Private Function getnum(u As String) As Long
'    If Len(u) <> 1 Then throw Exps.InvalidArgumentValueException
'    Dim b As Long
'    b = Asc(u)
'    b = b - 65
'    b = b + 15
'    If b >= 26 Then b = b - 6
'    If b < 0 Then throw Exps.InvalidArgumentValueException
'    getnum = b
'End Function

Public Function RegisterationState() As Boolean
    RegisterationState = LicensesCount > 0
End Function
Public Function LicensedFor(ByVal ProductType As ProductType) As Boolean

End Function


Public Function RegisterLibraryLicense(ByVal License As String, ByVal Info As String) As Boolean
    Dim PType As ProductType
    If ValidateLicense(License, Info, PType) Then
        ReDim Preserve Licenses(LicensesCount)
        Licenses(LicensesCount).Info = Info
        Licenses(LicensesCount).License = License
        Licenses(LicensesCount).keyCount = Len(License)
        Licenses(LicensesCount).Key = StringToByteArray(License)
        Licenses(LicensesCount).PType = PType
        LicensesCount = LicensesCount + 1
        RegisterLibraryLicense = True
    End If
End Function

'00 :
'01 :
'02 :
'03 :
'04 :
'05 :
'06 :
'07 : License Length.
'08 :
'09 :
'10 :
'11 :
'12 :
'13 :
'dyn:
'n-1: MintAPI Major Version
Public Function GenerateLicense(ByVal Info As String) As String
'    If Len(Info) < 20 Then throw Exps.InvalidArgumentValueException("Info length must be more than 20 ,ex(spaces are important,[] mean optional):MyCompany ,MyName ,Domain ex:com.google.app[ ,Email ,Phone ,...]")
'    Dim ln As Long
'    ln = MINCOUNT + (Rnd * (MAXCOUNT - MINCOUNT))
'    Dim lic() As Byte
'    ReDim lic(ln - 1)
'    lic(7) = getasc(ln)
'    lic(ln - 1) = (Asc(CStr(10 Mod App.Major)) * 2) + 3   'Version =
'    Dim k0 As Byte

End Function
Public Function ValidateLicense(ByVal License As String, ByVal Info As String, ByVal outType As ProductType) As Boolean
    ValidateLicense = True
End Function
