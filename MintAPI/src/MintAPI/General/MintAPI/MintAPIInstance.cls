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
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"Version"
Attribute VB_Ext_KEY = "Member1" ,"Path"
Attribute VB_Ext_KEY = "Member2" ,"Debugger"
Attribute VB_Ext_KEY = "Member3" ,"Module"
'@PROJECT_LICENSE

''<class name="MintAPIInstance" static tags="immutable">
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
Const CLASSID As String = "MintAPIInstance" '#DO NOT REMOVE IT
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

Public Property Get Module() As Module
    
End Property

Public Function GetTestingInstance() As Int64
Attribute GetTestingInstance.VB_MemberFlags = "40"
    Dim MWI As Object 'IMintAPIWrapperInstance

    Dim MWIPtr As Long
    MWIPtr = ObjPtr(MWI)
    GetTestingInstance.HigherValue = MWIPtr And &HFFFF 'FFFF
    GetTestingInstance.LowerValue = mHelper.ShiftRight(MWIPtr And &HFFFF0000, &H10)
End Function

Public Sub About(Optional ByVal Modal As Boolean = False)
Attribute About.VB_MemberFlags = "40"
    
End Sub

Public Property Get ThreadID() As Long
    ThreadID = App.ThreadID
End Property
Public Property Get LogPath() As String
    LogPath = App.LogPath
End Property

Public Property Get DllPath() As String
    Dim StrName As String
    StrName = App.EXEName
    DllPath = Path.Combine(App.Path, StrName)
End Property
Public Property Get Directory() As String
    Directory = App.Path
End Property

Public Property Get VersionInfo() As FileVersionInfo
    
End Property

'Public Property Get MintAPISharedPath() As String
'    MintAPISharedPath = Path.Combine(Path.GetSpecialFolder(sfLocalApplicationData), "MintAPI", "Shared")
'End Property


Public Function GetMintAPIVariable(ByVal Name As String, Optional DefaultValue) As ByteArray
'    On Error GoTo Err_Handler
'    Dim varBuffer() As Byte
'    varBuffer = mint_config.ReadMintAPIVariable(Name, True)
'    Set GetMintAPIVariable = ByteArray(varBuffer)
'    Exit Function
'Err_Handler:
'    Set GetMintAPIVariable = ByteArray(DefaultValue)
End Function
Public Sub SetMintAPIVariable(ByVal Name As String, ByVal Value As ByteArray)
    'Call modMain
End Sub
Public Function CheckMintAPIVariable(ByVal Name As String, ByRef outRetVal As ByteArray) As Boolean
    
End Function

Public Sub RunMintAPIHost(CommandArguments As String, Optional ByVal Stay As Boolean = False)
    
End Sub


Public Sub OptimizeForSingleThreadedApplication()
    
End Sub
Public Sub VirtualizeDesignMode()
    Call mint_application.VirtualizeDesignMode
End Sub

'*********************************************
'*********************************************
'
'</section>
'---------------------------------------------
'---------------------------------------------
