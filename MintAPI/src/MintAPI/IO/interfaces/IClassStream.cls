VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IClassStream"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'@PROJECT_LICENSE

''<class name="IClassStream" interface tags="">
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
Const CLASSID As String = "IClassStream" '#DO NOT REMOVE IT
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
Public Enum StreamDirection
    sdInStream = 1
    sdOutStream = 2
    sdBoth = sdInStream Or sdOutStream
End Enum
'
'</section>
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section Interface Body Members>
'
'*********************************************
'*********************************************

Public Function GetState(ByVal Dir As StreamDirection) As Boolean: End Function

Public Sub SetOutState(Optional ByVal State As Boolean = True, Optional Reserved): End Sub
Public Sub OutStream(ByVal Buffer As ClassStreamBuffer, Optional ByVal Length As Long = -1): End Sub

Public Sub SetInState(Optional ByVal State As Boolean = True, Optional Reserved): End Sub
Public Sub InStream(ByVal Buffer As ClassStreamBuffer, Optional ByVal Length As Long = -1): End Sub

Public Function ReadBytes(ByVal Length As Long, Optional ByVal StartAt As Long = 0) As Byte(): End Function
Public Function Available(ByVal Direction As StreamDirection) As Long: End Function

Public Sub OpenStream(Optional ByVal Direction As StreamDirection = StreamDirection.sdBoth): End Sub
Public Sub CloseStream(Optional ByVal Direction As StreamDirection = StreamDirection.sdBoth): End Sub

'*********************************************
'*********************************************
'
'</section>
'---------------------------------------------
'---------------------------------------------

