VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MintAPIInstance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"Version"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member1" ,"Path"
Attribute VB_Ext_KEY = "Member2" ,"Debugger"
Option Explicit


Public Function Reserved() As Long
    
End Function

Public Sub About(Optional Modal As Boolean = False)
    Dim A As New frmAbout
    Call A.Show(IIf(Modal, 1, 0))
End Sub
Public Property Get Copyright() As String
    Copyright = Mtr("Copyright (C) 2012 by Ali Mousavi Kherad | alimousavikherad@gmail.com")
End Property
Public Property Get ProductName() As String
    ProductName = App.ProductName
End Property
Public Property Get FileDescription() As String
    FileDescription = App.FileDescription
End Property
Public Property Get Comments() As String
    Comments = App.Comments
End Property
Public Property Get HelpFile() As String
    HelpFile = App.HelpFile
End Property
Public Property Get Handle() As Long
    Handle = App.hInstance
End Property
Public Property Get LegalTrademarks() As String
    LegalTrademarks = App.LegalTrademarks
End Property
Public Property Get LegalCopyright() As String
    LegalCopyright = App.LegalCopyright
End Property
Public Property Get ThreadID() As String
    ThreadID = App.ThreadID
End Property
Public Property Get LogPath() As String
    LogPath = App.LogPath
End Property

Public Function Debugger() As Debugger
    Set Debugger = mint_constructor.Debugger
End Function

Public Property Get DllPath() As String
    Dim StrName As String
    StrName = App.EXEName
    DllPath = Directory.ConcatPath(App.Path, StrName)
End Property
Public Property Get Path() As String
    Path = App.Path
End Property

Public Property Get VersionString() As String
    VersionString = APP_VERSIONSTRING
End Property
Public Property Get Version() As Version
    Set Version = MintAPI.Version.Create(App.Major, App.Minor, App.Revision, APP_RELEASEYEAR)
End Property
Public Function CheckVersion(ByVal MajorVersion As Long, Optional ByVal MinorVersion As Long = -1, Optional ByVal Revision As Long = -1, Optional ByVal CheckVersionTag As String = "", Optional ByVal CheckOneOfVersionTags As String = "") As Boolean
    If MajorVersion > App.Major Then Exit Function
    If MinorVersion <> -1 Then _
        If MinorVersion > App.Minor Then Exit Function
    If Revision <> -1 Then _
        If Revision > App.Revision Then Exit Function

    If (Trim(CheckVersionTag) = "") And (Trim(CheckOneOfVersionTags) = "") Then
        CheckVersion = True
    Else
        'CheckVersion = CheckVersionTag(CheckVersionTag, CheckOneOfVersionTags)
    End If
End Function
Public Property Get Int32Version() As Long
    Dim maj_Version As Long
    Dim min_Version As Long
    Dim rev_Version As Long
    maj_Version = tApplication.VersionMajor
    min_Version = tApplication.VersionMinor
    rev_Version = tApplication.VersionRevision
    If maj_Version > 255 Then GoTo outOfRangeExp_Err
    If min_Version > 255 Then GoTo outOfRangeExp_Err
    If rev_Version > 65535 Then GoTo outOfRangeExp_Err
    maj_Version = ShiftLeft(maj_Version And &HFF, 24)
    min_Version = ShiftLeft(min_Version And &HFF, 16)
    rev_Version = rev_Version And &HFFFF
    Version = (maj_Version + min_Version + rev_Version)
    Exit Property
outOfRangeExp_Err:
    Version = 0
End Property
Public Function CheckVersionTag(ByVal StrVersionTags As String, Optional ByVal StrConditionalVersionTags As String = "") As Boolean
    Dim strVersionTagsTRIMMED As String
    Dim strConditionalVersionTagsTRIMMED As String
    Dim APP_VERSIONTAGSTRIMMED As String
    strVersionTagsTRIMMED = Trim(StrVersionTags)
    strConditionalVersionTagsTRIMMED = Trim(StrConditionalVersionTags)
    If (strVersionTagsTRIMMED = "") And (strConditionalVersionTagsTRIMMED = "") Then
        APP_VERSIONTAGSTRIMMED = Trim(APP_VERSIONTAGS)
        If APP_VERSIONTAGS = "" Then
            CheckVersionTag = True
        Else
            CheckVersionTag = False
        End If
        Exit Function
    End If
    Dim str1() As String
    Dim str2() As String
    Dim sVersionTags() As String
    APP_VERSIONTAGSTRIMMED = Trim(APP_VERSIONTAGS)
    str1 = Split(Trim(StrVersionTags), " ")
    str2 = Split(Trim(StrConditionalVersionTags), " ")
    sVersionTags = Split(APP_VERSIONTAGSTRIMMED, " ")
    Dim str1Count As Long, str2Count As Long, sVTCount As Long
    str1Count = ArraySize(str1)
    str2Count = ArraySize(str2)
    sVTCount = ArraySize(sVersionTags)

    If str1Count <> sVTCount Then _
        If str1Count > 0 Then Exit Function

    Dim i As Long, j As Long, str2Found As Long
    For i = 0 To str1Count - 1
        For j = 0 To sVTCount - 1
            If str1(i) <> sVersionTags(j) Then
                CheckVersionTag = False
                Exit Function
            End If
        Next
    Next
    For i = 0 To str2Count - 1
        For j = 0 To sVTCount - 1
            If str2(i) = sVersionTags(j) Then
                CheckVersionTag = True
                str2Found = str2Found + 1
                Exit Function
            End If
        Next
    Next
    If str1Count <= 0 Then
        If str2Found > 0 Then
            CheckVersionTag = True
        Else
            CheckVersionTag = False
        End If
    Else
        CheckVersionTag = True
    End If
End Function

Public Function RegisterLibraryLicense(ByVal License As String, ByVal Info As String) As Boolean
    RegisterLibraryLicense = licensing.RegisterLibraryLicense(License, Info)
End Function
