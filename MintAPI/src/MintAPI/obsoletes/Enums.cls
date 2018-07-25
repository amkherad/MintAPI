VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Enums"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'@PROJECT_LICENSE

Option Base 0
Option Explicit
Const CLASSID As String = "Enums"
Const CLASSALIAS As String = APPLICATIONDOMAIN & "." & CLASSID
Const CLASSVERSIONSTRING As String = "0.1"
Const CLASSALIASSTRING As String = CLASSALIAS & "/" & CLASSVERSIONSTRING


Implements IAliasable

Friend Sub Initialize(Optional Arguments)
    
End Sub

'
'Public Function EnumSocketTypeToString(SockType As SocketType, Optional UseIdetifierKey As Boolean = True) As String
'    Select Case SockType
'        Case SocketType.stDataGram
'            EnumSocketTypeToString = "DataGram"
'        Case SocketType.stStream
'            EnumSocketTypeToString = "Stream"
'        Case SocketType.stRAW
'            EnumSocketTypeToString = "RAW"
'        Case SocketType.stReliablyDeliveredMessage
'            EnumSocketTypeToString = "ReliablyDeliveredMessage"
'        Case SocketType.stSequencedPacket
'            EnumSocketTypeToString = "SequencedPacket"
'        Case Else
'            EnumSocketTypeToString = "Unknown"
'    End Select
'
'    If UseIdetifierKey Then _
'        If EnumSocketTypeToString <> "" Then _
'            EnumSocketTypeToString = "st" & EnumSocketTypeToString
'End Function
Public Function EnumCompareValueToString(ByVal CompValue As CompareValue, Optional UseIdetifierKey As Boolean = True) As String
    If CompValue > 0 Then
        EnumCompareValueToString = "Greater"
    ElseIf CompValue < 0 Then
        EnumCompareValueToString = "Less"
    Else
        EnumCompareValueToString = "Equal"
    End If
    
    If UseIdetifierKey Then _
        If EnumCompareValueToString <> "" Then _
            EnumCompareValueToString = "cv" & EnumCompareValueToString
End Function
Public Function EnumTristateValuesToString(ByVal TriValue As TristateValues, Optional UseIdetifierKey As Boolean = True) As String
    If TriValue > 0 Then
        EnumTristateValuesToString = "More"
    ElseIf TriValue < 0 Then
        EnumTristateValuesToString = "Less"
    Else
        EnumTristateValuesToString = "Equal"
    End If
    
    If UseIdetifierKey Then _
        If EnumTristateValuesToString <> "" Then _
            EnumTristateValuesToString = "tv" & EnumTristateValuesToString
End Function

Private Property Get IAliasable_Alias() As String
    IAliasable_Alias = CLASSALIASSTRING
End Property
