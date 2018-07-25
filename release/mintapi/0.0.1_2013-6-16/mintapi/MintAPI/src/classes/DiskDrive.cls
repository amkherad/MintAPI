VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DiskDrive"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'@PROJECT_LICENSE

Option Base 0
Option Explicit
Const CLASSID As String = "DiskDrive"
Const CLASSALIAS As String = APPLICATIONDOMAIN & "." & CLASSID
Const CLASSVERSIONSTRING As String = "0.1"
Const CLASSALIASSTRING As String = CLASSALIAS & "/" & CLASSVERSIONSTRING


Implements IAliasable


Dim DL As String

Friend Sub Initialize(Optional DriveLetter)
    If Not IsMissing(DriveLetter) Then DL = DriveLetter
End Sub

Public Sub MountFolder(DriveLetter As String, FolderPath As String, Optional VolumeName As String)
    Call kernelMethods.MountFolder(DriveLetter, FolderPath, VolumeName)
End Sub
Public Sub UnmountFolder(DriveLetter As String, FolderPath As String)
    Call kernelMethods.UnmountFolder(DriveLetter, FolderPath)
End Sub
Public Sub EjectDrive(Optional OptionalDriveLetter)
    Dim VL As String
    VL = IIf(IsMissing(OptionalDriveLetter), DL, CStr(OptionalDriveLetter))
    Call kernelMethods.EjectDrive(VL)
End Sub
Public Function Drives() As String()
    Dim drvStr As String
    drvStr = kernelMethods.GetLogicalDrivesString
    Dim splitStr() As String
    splitStr = Split(drvStr, " ")
    Drives = splitStr
End Function
Public Function DriveType(Optional OptionalDriveLetter) As DriveType
    Dim VL As String
    VL = IIf(IsMissing(OptionalDriveLetter), DL, CStr(OptionalDriveLetter))
    If VL = "" Then throw InvalidArgumentValueException("Invalid Drive Letter.")
    DriveType = kernelMethods.GetDriveType(VL)
End Function
Public Function VolumeName(Optional OptionalDriveLetter) As String
    Dim VL As String
    VL = IIf(IsMissing(OptionalDriveLetter), DL, CStr(OptionalDriveLetter))
    If VL = "" Then throw InvalidArgumentValueException("Invalid Drive Letter.")
    VolumeName = kernelMethods.VolumeName(VL)
End Function
Public Sub SetVolumeName(VolumeName As String, Optional OptionalDriveLetter)
    Dim VL As String
    VL = IIf(IsMissing(OptionalDriveLetter), DL, CStr(OptionalDriveLetter))
    If VL = "" Then throw InvalidArgumentValueException("Invalid Drive Letter.")
    Call kernelMethods.SetVolumeName(VL, VolumeName)
End Sub
Public Sub SetNTFSPermissions()
    
End Sub

Private Property Get IAliasable_Alias() As String
    IAliasable_Alias = CLASSALIASSTRING
End Property
