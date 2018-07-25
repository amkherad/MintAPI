VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IMemoryBlock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'@PROJECT_LICENSE

''<class name="IMemoryBlock" interface tags="">
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
Const CLASSID As String = "IMemoryBlock" '#DO NOT REMOVE IT
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
Public Enum MemoryAccessPermissions
    mapNoAccess = PAGE_NOACCESS
    mapReadOnly = PAGE_READONLY
    mapReadWrite = PAGE_READWRITE
    mapWriteCopy = PAGE_WRITECOPY
    mapExecute = PAGE_EXECUTE
    mapExecuteRead = PAGE_EXECUTE_READ
    mapExecuteWriteCopy = PAGE_EXECUTE_WRITECOPY
    mapExecuteReadWrite = PAGE_EXECUTE_READWRITE
    mapGuard = PAGE_GUARD
    mapNoCache = PAGE_NOCACHE
    mapFullControl = PAGE_EXECUTE_READWRITE
    
    mapDefault = mapReadWrite
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

Public Property Get Reference() As Long: End Property

Public Property Get IsLimited() As Boolean: End Property
Public Property Get Limit() As Long: End Property

Public Property Get Access() As MemoryAccessPermissions: End Property

Public Property Get IsResizable() As Boolean: End Property
Public Sub Resize(ByVal NewSize As Long): End Sub

'*********************************************
'*********************************************
'
'</section>
'---------------------------------------------
'---------------------------------------------

