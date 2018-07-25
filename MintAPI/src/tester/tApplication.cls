VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "tApplication"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'----------------------------------------------
'  MintAPI by Ali Mousavi Kherad
'  - alimousavikherad@gmail.com
'----------------------------------------------

Option Explicit
Implements IApplication

Const VERSIONSTRING As String = "0.0.1.2013 blackbox - advanced"

Const COMPILEYEAR As String = 2013
Const COMPILEYEARS As String = "2013"
Const COMPILEDATE As String = "0/0/2013"
Const COMPILERINFO As String = "VB6 Advanced"
Const COMPILEBY As String = "Compiled by | support@test.com"


Private Property Get IApplication_Name() As String: IApplication_Name = App.Title: End Property

Private Sub IApplication_Main(Arguments As Variant)
    'code here.
    'Main Application Entry...
    'Signals.Connect "ctrl_c", Console, Slot("terminate", Application)
    'Dim Arr As Variant, Arr1() As Variant, Arr2(0) As Variant
    'Dim Arr As Variant, Arr1() As Variant, Arr2(0) As Variant
    'Dim API As Object, l As Object, l2 As Object, Q As New Queue, S As Object, J
    'Set API = CreateObject("MintAPI.API")
    'Set l = CreateObject("MintAPI.List")
    
    MsgBox mString.Trim("ttAli Mousavi Kheradt", "t")
    
    'Q.Enqueue "Alisdf"
    'Q.Enqueue "Alisdf2"
    'Q.Enqueue "Alisdf3"
    
    
    'MsgBox "Queue:" & Q.Dequeue
    'MsgBox "Queue:" & Q.Dequeue
    'MsgBox "Queue:" & Q.Dequeue
''
End Sub
Public Sub ShowArray(ByVal Name As String, A)
    Dim Str As String, Count As Long
    On Error Resume Next
    Count = UBound(A) - LBound(A) + 1
    Dim V
    For Each V In A
        Str = Str & CStr(V) & vbNewLine
    Next
    MsgBox Name & "[" & Count & "]" & vbNewLine & Str
End Sub
Private Sub IApplication_Starting(ByVal E As MintAPI.EventArgs)
Call Debugger.StartLog
    Debugger.Log "Starting..."
End Sub
Private Sub IApplication_Started(ByVal E As MintAPI.EventArgs)
    Debugger.Log "Started..."
    'code here.
End Sub
Private Sub IApplication_Restart(ByVal E As MintAPI.EventArgs)
    Debugger.Log "Restart..."
    'code here.
End Sub
Private Sub IApplication_Terminating(ByVal E As MintAPI.CancelEventArgs)
    Debugger.Log "Terminating ..."
    'E.Cancel = True 'May cancel the termination operation.
    'code here.
End Sub
Private Sub IApplication_Terminated(ByVal E As MintAPI.EventArgs)
    Debugger.Log "Terminated."
    'code here.
    '-------------------
    Call Debugger.EndLog
End Sub
Private Sub IApplication_Debugger(ByVal E As MintAPI.DebuggerNotificationEventArgs)
    '#WARNING: Don't use Debugger.XXX here...
    
    'code here.
End Sub
Private Sub IApplication_Error(ByVal E As MintAPI.ExceptionOccuredEventArgs)
    '#WARNING: Don't use Debugger.XXX here...
    
    'code here.
End Sub
Private Sub IApplication_ApplicationInputMethodChanged(ByVal E As MintAPI.EventArgs)
    'code here.
End Sub
Private Sub IApplication_ApplicationNetworkStateChanged(ByVal E As MintAPI.EventArgs)
    
    'code here.
End Sub
Private Sub IApplication_ApplicationConfigurationChanged(ByVal E As MintAPI.EventArgs)
    Debugger.Log "Current Configuration:", Configuration.ConfigurationType
    'code here.
End Sub
Private Sub IApplication_ApplicationLanguageChanged(ByVal E As MintAPI.EventArgs)
    Debugger.Log "Current Language:", Application.CurrentLanguage.Name
    'code here.
End Sub
Private Sub IApplication_Message(ByVal E As MintAPI.ApplicationMessageEventArgs)
    Debugger.Log "Message recieved..."
    'code here.
End Sub

'Indicates the unique application alias.
Private Property Get IApplication_AppAlias() As String
    IApplication_AppAlias = "com.mintapi.testapp"
End Property
Private Property Get IApplication_Version() As Version
    Set IApplication_Version = Version.Create(App.Major, App.Minor, COMPILEYEAR, App.Revision)
End Property
Private Property Get IApplication_CompanyName() As String
    IApplication_CompanyName = App.CompanyName
End Property
Private Property Get IApplication_ProductName() As String
    IApplication_ProductName = App.ProductName
End Property
Private Property Get IApplication_LegalCopyright() As String
    IApplication_LegalCopyright = App.LegalCopyright
End Property
Private Property Get IApplication_ProductCode50Chars() As String: End Property
Private Property Get IApplication_ProductCode() As String: End Property
Private Property Get IApplication_ApplicationGuid() As MintAPI.Guid: End Property
Private Property Get IApplication_ApplicationLinks() As ApplicationLinks
    IApplication_ApplicationLinks.Website = "http://www.mintapi.com/"
    IApplication_ApplicationLinks.DocumentationURL = "http://www.mintapi.com/?p=docs"
    IApplication_ApplicationLinks.HelpURL = "http://www.mintapi.com/?p=help"
    IApplication_ApplicationLinks.ServiceURL = "http://www.mintapi.com/?p=service&args=[QUERY_VALUE_ARGS]"
    IApplication_ApplicationLinks.SupportURL = "http://www.mintapi.com/?p=support&q=[QUERY_VALUE_QUESTION]"
    IApplication_ApplicationLinks.UpdateURL = "http://www.mintapi.com/?p=update&q=[QUERY_VALUE]"
End Property
