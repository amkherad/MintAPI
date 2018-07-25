VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MCIPlayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'@PROJECT_LICENSE

''<class name="BasicPlayer" static tags="">
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
Const CLASSID As String = "BasicPlayer" '#DO NOT REMOVE IT
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
Public Enum CDDoorStatus
    cddsOpen
    cddsClose
End Enum
'
'</section>
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section Creators>
'
Public Function CDAudio(Optional ByVal hWnd As Long = 0, Optional ByVal Flags As Long) As MCIPlayerMedia
    Dim ID As Long, Alias As String
    Alias = "cdaudio"
    
    'If Alias = "" Then Alias = mString.GetRandomString(ctAlphabetic, 10, 20)
    
    Call mciOpenFile("", Alias, True, MCI_DEVTYPE_CD_AUDIO, hWnd, Flags)
    
    Set CDAudio = New MCIPlayerMedia
    Call CDAudio.Constructor0("", ID, Alias, hWnd, True)
End Function
Public Function OpenFromID(ByVal ID As String, ByVal IsOpen As Boolean, Optional ByVal Alias As String = "") As MCIPlayerMedia
    Set OpenFromID = New MCIPlayerMedia
    Call OpenFromID.Constructor1(ID, Alias, IsOpen)
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

Public Function OpenMediaFile(ByVal Path As String, Optional ByVal Alias As String, Optional ByVal Sharable As Boolean = True, Optional ByVal hWnd As Long = 0) As MCIPlayerMedia
    Dim ID As Long
    
    If Alias = "" Then Alias = mString.GetRandomString(ctAlphabetic, 10, 20)
    
    Call mciOpenFile(Path, Alias, Sharable, MCI_ALL_DEVICE_ID, hWnd)
    
    Set OpenMediaFile = New MCIPlayerMedia
    Call OpenMediaFile.Constructor0(Path, ID, Alias, hWnd, True)
End Function

Public Sub SetCDDoorStatus(ByVal Status As CDDoorStatus)
    
End Sub

'*********************************************
'*********************************************
'
'</section>
'---------------------------------------------
'---------------------------------------------
