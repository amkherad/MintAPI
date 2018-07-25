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

Private Type LICENSE
    keyCount As Long
    Key() As Byte
    Info As String
    LICENSE As String
    PType As ProductType
End Type
Dim licenses() As LICENSE
Dim licensesCount As Long


Public Sub Initialize()
    Call RegisterLibraryLicense("0000000000000000000000000", "MintAPI Learning License Version.")
    
End Sub
Public Sub Dispose(Optional ByVal Force As Boolean = False)
    
End Sub

'Private Function getch(b As Byte) As String
'    Dim u As String * 1
'    b = b - MINCOUNT
'    If b >= 26 Then b = b + 6
'    If b > 66 Then throw InvalidArgumentValueException
'    u = Chr(65 + (b))
'    getch = u
'End Function
Private Function getasc(ByVal b As Byte) As Byte
    b = b - MINCOUNT
    If b >= 26 Then b = b + 6
    If b > 66 Then throw InvalidArgumentValueException
    getasc = 65 + (b)
End Function
Private Function getchr(ByVal b As Long) As Long
    If b > 66 Then throw InvalidArgumentValueException
    b = b - 65
    b = b + 15
    If b >= 26 Then b = b - 6
    If b < 0 Then throw InvalidArgumentValueException
    getchr = b
End Function
'Private Function getnum(u As String) As Long
'    If Len(u) <> 1 Then throw InvalidArgumentValueException
'    Dim b As Long
'    b = Asc(u)
'    b = b - 65
'    b = b + 15
'    If b >= 26 Then b = b - 6
'    If b < 0 Then throw InvalidArgumentValueException
'    getnum = b
'End Function

Public Function RegisterationState() As Boolean
    RegisterationState = licensesCount > 0
End Function
Public Function LicensedFor(ProductType As ProductType) As Boolean

End Function


Public Function RegisterLibraryLicense(LICENSE As String, Info As String) As Boolean
    Dim PType As ProductType
    If ValidateLicense(LICENSE, Info, PType) Then
        ReDim Preserve licenses(licensesCount)
        licenses(licensesCount).Info = Info
        licenses(licensesCount).LICENSE = LICENSE
        licenses(licensesCount).keyCount = Len(LICENSE)
        licenses(licensesCount).Key = StringToByteArray(LICENSE)
        licenses(licensesCount).PType = PType
        licensesCount = licensesCount + 1
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
Public Function GenerateLicense(Info As String) As String
    If Len(Info) < 20 Then throw InvalidArgumentValueException("Info length must be more than 20 ,ex(spaces are important,[] mean optional):MyCompany ,MyName ,Domain ex:com.google.app[ ,Email ,Phone ,...]")
    Dim ln As Long
    ln = MINCOUNT + (Rnd * (MAXCOUNT - MINCOUNT))
    Dim lic() As Byte
    ReDim lic(ln - 1)
    lic(7) = getasc(ln)
    lic(ln - 1) = (Asc(CStr(10 Mod App.Major)) * 2) + 3   'Version =
    Dim k0 As Byte

End Function
Public Function ValidateLicense(LICENSE As String, Info As String, outType As ProductType) As Boolean
    ValidateLicense = True
End Function
