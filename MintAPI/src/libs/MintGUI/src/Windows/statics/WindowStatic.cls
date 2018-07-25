VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WindowStatic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'@PROJECT_LICENSE

''<class name="WindowStatic" static tags="">
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
Const CLASSID As String = "WindowStatic" '#DO NOT REMOVE IT
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
Public Function FromHandle(ByVal WindowHandle As Long) As Window
Attribute FromHandle.VB_UserMemId = 0
    Set FromHandle = New Window
    Call FromHandle.Constructor0(WindowHandle)
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

Public Property Get Desktop() As Window: Set Desktop = FromHandle(API_GetDesktopWindow): End Property
Public Property Get DesktopHandle() As Window: DesktopHandle = API_GetDesktopWindow: End Property


Public Property Get TopWindow() As Window
    Set TopWindow = FromHandle(API_GetTopWindow(vbNullPtr))
End Property

Public Property Get StateChangeAnimation() As Boolean
    Dim Info As API_ANIMATIONINFO
    Call API_SystemParametersInfo(SPI_GETANIMATION, Info.cbSize, Info, 0)
    StateChangeAnimation Info.iMinAnimate <> 0
End Property
Public Property Let StateChangeAnimation(ByVal Value As Boolean)
    Dim Info As API_ANIMATIONINFO
    
    'ANIMATIONINFO info = new ANIMATIONINFO(value ? 1 : 0)
    Call SystemParametersInfo(SPI_SETANIMATION, Info.cbSize, Info, 0)
End Property

'*********************************************
'*********************************************
'
'</section>
'---------------------------------------------
'---------------------------------------------
