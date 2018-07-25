VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Path"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'@PROJECT_LICENSE

''<class name="Path" static tags="">
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
Const CLASSID As String = "Path" '#DO NOT REMOVE IT
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

'
'</section>
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section Creators>
'
'Public Function Create(ByVal Args As Long) As Object
'
'End Function
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

''<default/>
Public Function Validate(ByVal Path As String) As String
Attribute Validate.VB_UserMemId = 0
    
End Function

'by Kelly Ethridge , VBCorLib:Path #2004
''<summary>Returns the maximum length for a path.</summary>
''<retval>Tha maximum length allowed for a path.</retval>
''<remarks>Returns 32767 on Windows NT machines, 260 for Win9x machines.</remarks>
Public Property Get MaxPathLength() As Long
    If Environment.IsUnicode Then
        MaxPathLength = MAX_PATH_W
    Else
        MaxPathLength = MAX_PATH_A
    End If
End Property

''<summary>Returns the character code for the alternative directory separator character.</summary>
''<retval>The character code of the alternative directory separator (47).</retval>
''<remarks>The alternate character is usually found on systems such as Unix.</remarks>
Public Property Get AltDirectorySeparatorChar() As Integer
    AltDirectorySeparatorChar = vbForwardSlash
End Property
''<summary>Returns the string for the alternative directory separator character.</summary>
''<retval>The string of the alternative directory separator "/".</retval>
''<remarks>The alternate string is usually found on systems such as Unix.</remarks>
Public Property Get AltDirectorySeparatorString() As String
    AltDirectorySeparatorString = vbForwardSlashS
End Property

''<summary>Returns the directory separator character for Windows.</summary>
''<retval>The directory separator character for Windows (92).</retval>
Public Property Get DirectorySeparatorChar() As Integer
    DirectorySeparatorChar = vbBackSlash
End Property

''<summary>Returns the directory separator string for Windows.</summary>
''<retval>The directory separator string for windows "\".</retval>
Public Property Get DirectorySeparatorString() As String
    DirectorySeparatorString = vbBackSlashS
End Property

''<summary>Returns an array of characters that are invalid in a Windows path string.</summary>
''<retval>An Integer array representing the character code of the the invalid characters.</retval>
Public Property Get InvalidPathChars() As Integer()
    'InvalidPathChars = mInvalidPathChars
End Property

''<summary>The character code used to separate strings in an environment variable.</summary>
''<retval>The environment variable string separator character code (59).</retval>
Public Property Get PathSeparator() As Integer
    PathSeparator = vbSemiColon
End Property

''<summary>The string used to seperate strings in an environment variable.</summary>
''<retval>The environment variable string separator ";".</retval>
Public Property Get PathSeparatorString() As String
    PathSeparatorString = vbSemiColonS
End Property

''<summary>The character code used to separate volume labels from path information.</summary>
''<retval>The character code for the volume separator (58).</retval>
Public Property Get VolumeSeparatorChar() As Integer
    VolumeSeparatorChar = vbColon
End Property

''<summary>Returns the string version of the volume separator.</summary>
''<retval>The string version of the volume separator ":".</retval>
Public Property Get VolumeSeparatorString() As String
    VolumeSeparatorString = vbColonS
End Property

Public Function ChangeExtension(ByVal Path As String, ByVal Extension As String) As String
    If LenB(Path) = 0 Then Exit Function
    Dim DotIndex As Long, SlashIndex As Long
    Path = Replace$(Path, "/", "\")
    DotIndex = InStrRev(Path, ".", 1)
    SlashIndex = InStrRev(Path, "\", 1)
    
    If DotIndex > SlashIndex Then
        
    Else
        ChangeExtension = Combine(Path, Extension)
    End If
End Function

Public Function IsPathRooted(ByVal Path As String) As Boolean
    
End Function

Public Function HasExtension(ByVal Path As String) As Boolean
    
End Function

Public Function Combine(ParamArray Path() As Variant) As String
    Dim RetVal As String, P As Variant
    For Each P In Path
        RetVal = RetVal & "\" & mString.Stringify(P)
    Next
End Function

Public Function GetDirectoryName(ByVal Path As String) As String
    
End Function

Public Function GetExtension(ByVal Path As String) As String
    
End Function

Public Function GetFileName(ByVal Path As String) As String
    
End Function

Public Function GetFileNameWithoutExtension(ByVal Path As String) As String
    
End Function

Public Function GetFullPath(ByVal Path As String) As String
    
End Function

Public Function GetPathRoot(ByVal Path As String) As String
    
End Function

Public Function GetTempFileName() As String
    
End Function

Public Function GetTempPath() As String
    
End Function

Public Function GetInvalidPathChars() As Integer()
    
End Function

Public Function GetInvalidFileNameChars() As Integer()
    
End Function


'Public Function GetSpecialFolder(ByVal SpecialFolders As SpecialFolders) As String
'    Dim Result As Long
'    Dim IDL As API_ITEMIDLIST
'On Error GoTo CatchErr:
'    'Get the special folder
'    Result = API_SHGetSpecialFolderLocation(0, SpecialFolders, IDL)
'    If Result = 0 Then
'        'Create a buffer
'        Dim Path As String
'        Path = String(LARGELPSTR, Chr(0))
'        'Get the path from the IDList
'        Result = API_SHGetPathFromIDList(ByVal IDL.mkid.cb, ByVal Path$)
'        'Remove the unnecessary chr$(0)'s
'        GetSpecialFolder = Left$(Path, InStr(Path, Chr$(0)) - 1)
'        Call API_SHGetMalloc.Free(VarPtr(IDL))
'        Exit Function
'    End If
'    throw Exps.Exception("Unknown SpecialFolder value.")
'Exit Function
'CatchErr:
'    throw Exps.InvalidOperationException(Err.Description)
'End Function

'*********************************************
'*********************************************
'
'</section>
'---------------------------------------------
'---------------------------------------------
