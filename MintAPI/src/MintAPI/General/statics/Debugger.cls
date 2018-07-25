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
Attribute VB_Ext_KEY = "Member0" ,"Debugger"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Declare Function API_GetTickCount Lib "kernel32" Alias "GetTickCount" () As Long


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
    If FilterLevel < 0 Then throw Exps.NegativeArgumentException
    p_dbgFilterLevel = FilterLevel
End Sub
Public Sub Debugger(ByVal Message As String, Optional ByVal Level As Long = 1, Optional ByVal TargetApplication As IApplication)
    If Level < 1 Then Level = 1
    Call Log(Message, vbRed, TargetApplication)
    If Level <= p_dbgFilterLevel Then Exit Sub

End Sub
Public Sub Log(ByVal Message As String, Optional ByVal Color As Long = vbBlack, Optional ByVal TargetApplication As IApplication)
    If TargetApplication Is Nothing Then Set TargetApplication = AppInfo.TargetApplication
    If AppInfo.TargetApplication Is Nothing Then throw Exps.InvalidStatusException
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

Public Function ReadProcessMemory(ByVal Process As Process, ByVal BaseAddress As Long, ByVal Length As Long) As Memory
    'Set ReadProcessMemory = Process.ReadProcessMemory(BaseAddress, Length)
End Function
Public Function WriteProcessMemory(ByVal Process As Process, ByVal BaseAddress As Long, ByVal Length As Long) As Memory
    'Set ReadProcessMemory = Process.ReadProcessMemory(BaseAddress, Length)
End Function
