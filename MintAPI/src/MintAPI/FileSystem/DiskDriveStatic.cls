VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DiskDriveStatic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'@PROJECT_LICENSE

''<class name="DiskDriveStatic" static tags="">
''  <summary>
''    Some class description.
''  </summary>
''  <using>Some class alias.</using>
''  <idea>Me</idea>
''  <sample></sample>
''</class>

Option Base 0
Option Explicit

'=============================================
'=============================================
'=============================================
'<section Meta Data>
'
Const CLASSID As String = "DiskDriveStatic" '#DO NOT REMOVE IT
Const CLASSALIAS As String = APPLICATIONDOMAIN & "." & CLASSID '#DO NOT REMOVE IT
Const CLASSVERSIONSTRING As String = "0.1" '#DO NOT REMOVE IT
Const CLASSALIASSTRING As String = CLASSALIAS & "/" & CLASSVERSIONSTRING '#DO NOT REMOVE IT

'
'</section>
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section API/Type Declarations>
'
Private Const API_DDD_REMOVE_DEFINITION As Long = &H2

Private Const DRIVE_REMOVABLE = 2
Private Const DRIVE_FIXED = 3
Private Const DRIVE_REMOTE = 4
Private Const DRIVE_CDROM = 5
Private Const DRIVE_RAMDISK = 6
'
'</section>
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section Creators>
'
Public Function FromLetter(ByVal RootPath As String) As DiskDrive
Attribute FromLetter.VB_UserMemId = 0
    Set FromLetter = New DiskDrive
    Call FromLetter.Constructor0(RootPath)
End Function
'
'</section>
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section Class Public Body Members>
'
'*********************************************
'*********************************************

Public Function Drives() As Vector
    Dim dvsStr As String, V As Variant
    
    Dim Buf As String * SMALLLPSTR, bufSize As Long
    Buf = String$(SMALLLPSTR, " ")
    bufSize = API_GetLogicalDriveStrings(SMALLLPSTR, Buf)
    If bufSize = NO_VALUE Then throw Exps.IfError
    dvsStr = Trim$(Replace(Buf, Chr(0), " "))
    
    Dim dvsStrs() As String
    dvsStrs = Split(Buf, " ")
    
    Dim RetLst As New List, Drv As DiskDrive
    For Each V In dvsStrs
        Set Drv = New DiskDrive
        Call Drv.Constructor0(V)
        Call RetLst.Append(Drv)
    Next
    Set Drives = RetLst.ToVector
End Function

Friend Function ValidateRootPath(ByVal RootPath As String) As String
    RootPath = Trim(RootPath)
    Dim ln As Long
    ln = Len(RootPath)
    If ln <= 0 Then GoTo RaiseError
    If ln = 1 Then
        ' do nothing
    ElseIf ln = 2 Then
        If Not (RootPath Like "?:") Then GoTo RaiseError
    ElseIf ln = 3 Then
        If Not (RootPath Like "?:[/,\,|]") Then GoTo RaiseError
    Else
        If Not (Left(RootPath, 3) Like "?:[/,\,|]") Then GoTo RaiseError
    End If
    ValidateRootPath = Left(RootPath, 1)
    Exit Function
RaiseError:
    throw Exps.Exception("Invalid Drive Letter.")
End Function


Public Function Mount(ByVal DriveLetter As String, ByVal Provider As IProvider) As DiskDrive
    
End Function


Public Function MapFolderAsDrive(ByVal DriveLetter As String, _
        ByVal Path As String, Optional ByVal VolumeName As String, _
        Optional ByVal RetainUntillRestart As Boolean = False) As DiskDrive
    If Len(DriveLetter) <> 1 Then throw Exps.InvalidArgumentException("DriveLetter")
    If Not mString.IsAlphabetic(DriveLetter) Then throw Exps.InvalidArgumentException("DriveLetter")
    
    Dim lDriveType As Long, sPath As String
    'Remove any white spaces ...
    'sDriveLetter = ValidateDriveLetter(sDriveLetter)
    'Check if specified Folder Path is correct & exists ...
    If Not Directory.Exists(Path) Then _
        throw Exps.Exception("Specified mount path is wrong or does not point to a valid Windows Folder item.")
    'DefineDosDevice requires ':' at the end of drive letter ...
    sPath = DriveLetter & ":"
    'Oops ! For GetDriveType, we need to append a \ to Drive Letter. Let us check if specified Drive Letter is available ...
    lDriveType = API_GetDriveType(sPath & "\")
    'Only Unknown type of Drive letters are allowed to use for Virtual Mount for obvious reasons ...
    Select Case lDriveType
        Case DRIVE_CDROM, DRIVE_FIXED, DRIVE_RAMDISK, DRIVE_REMOVABLE, DRIVE_REMOTE:
                throw Exps.Exception("Specified Drive letter is not available to mount virtual drive.")
        'Here it means that the Drive Letter is available for us to mount Virtual Drive ...
        Case Else
            If Not CBool(API_DefineDosDevice(0, sPath, Path)) Then _
                throw Exps.IfError
    End Select
    
    Set MapFolderAsDrive = New DiskDrive
    Call MapFolderAsDrive.Constructor0(DriveLetter, Not RetainUntillRestart)
End Function
Friend Sub MapFolderAsDriveReverse(ByVal DriveLetter As String, _
        ByVal Path As String)
    If Not CBool(API_DefineDosDevice(API_DDD_REMOVE_DEFINITION, DriveLetter, Path)) Then _
        throw Exps.IfError
End Sub
'*********************************************
'*********************************************
'
'</section>
'---------------------------------------------
'---------------------------------------------
