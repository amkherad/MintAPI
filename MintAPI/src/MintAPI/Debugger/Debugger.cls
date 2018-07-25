VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Debugger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'@PROJECT_LICENSE

''<class name="" static tags="">
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
Const CLASSID As String = "MintAPIStaticClass" '#DO NOT REMOVE IT
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
Public Enum DebuggerFormatConstants
    dfcNormalText
    dfcHTML
End Enum
Public Enum DebuggerDataShowMode
    dbgsAppend = 0
    dbgsClear = 1
    dbgsHighlightLast = &H10
    dbgsModal = &H4000
End Enum

Public Type DebuggerAttributes
    LogPath As String
    Level As Long
    Format As DebuggerFormatConstants
    DataShowMode As DebuggerDataShowMode
    Color As Long
End Type

Dim p_dbgFilterLevel As Long
Dim p_DbgInfo As DebuggerAttributes
Dim p_OFS As FileStream
Dim p_StopWatchValue As Long
'
'</section>
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section Creators>
'

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

Public Sub Breakpoint()
    Call Exps.RaiseSystemException(STATUS_BREAKPOINT, NO_VALUE)
End Sub
Public Sub SingleStep()
    Call Exps.RaiseSystemException(STATUS_SINGLE_STEP, NO_VALUE)
End Sub

Public Sub StartLog(Optional ByVal LogPath As String, Optional ByVal Format As DebuggerFormatConstants)
    'Set p_OFS = File
    
    p_DbgInfo.LogPath = LogPath
    p_DbgInfo.Format = Format
End Sub
Public Sub EndLog()
    
End Sub

Public Property Get IsDebuggerAttached() As Boolean
    
End Property

Public Sub FilterMessageLevels(Optional ByVal FilterLevel As Long = 2)
    If FilterLevel < 0 Then throw Exps.InvalidArgumentException
    p_dbgFilterLevel = FilterLevel
End Sub
Public Sub Log(ByVal Message As String, Optional ByVal Color As Long = vbBlack, Optional ByVal TargetApplication As IApplication)
    If TargetApplication Is Nothing Then Set TargetApplication = AppInfo.TargetApplication
    If AppInfo.TargetApplication Is Nothing Then throw Exps.InvalidOperationException
    'Format(NOW,"MM\DD\YYYY hh:mm:ss")
End Sub
Public Sub Dump(ByRef Expression As Variant)
    Call Console.WriteLine(CStr(Expression))
End Sub

Public Function Show(ByRef Data As Variant, Optional ByVal Mode As DebuggerDataShowMode = DebuggerDataShowMode.dbgsHighlightLast) As Boolean
    Call MsgBox(CStr(Data))
End Function

Public Function ShowFatalErrorAndExit(ByVal Caption As String, _
        ByVal Description As String, _
        ByVal Alias As String, _
        Optional ByVal tObj As Object = Nothing) As Boolean
    Call MsgBox(Replace(Description, "\n", vbNewLine), Title:=Caption)
    Call Process.EndProcess
End Function

Public Function ReadProcessMemory(ByVal Process As Process, ByVal BaseAddress As Long, ByVal Length As Long) As IMemoryBlock
    'Set ReadProcessMemory = Process.ReadProcessMemory(BaseAddress, Length)
End Function
Public Function WriteProcessMemory(ByVal Process As Process, ByVal BaseAddress As Long, ByVal Length As Long) As IMemoryBlock
    'Set ReadProcessMemory = Process.ReadProcessMemory(BaseAddress, Length)
End Function

'*********************************************
'*********************************************
'
'</section>
'---------------------------------------------
'---------------------------------------------
