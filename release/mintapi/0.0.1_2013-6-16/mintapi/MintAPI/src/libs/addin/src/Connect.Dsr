VERSION 5.00
Begin {AC0714F6-3D04-11D1-AE7D-00A0C90F26F4} Connect 
   ClientHeight    =   13440
   ClientLeft      =   1740
   ClientTop       =   1545
   ClientWidth     =   14295
   _ExtentX        =   25215
   _ExtentY        =   23707
   _Version        =   393216
   Description     =   "MintAPI Design Tools"
   DisplayName     =   "MintAPI Environment"
   AppName         =   "Visual Basic"
   AppVer          =   "Visual Basic 6.0"
   LoadName        =   "Startup"
   LoadBehavior    =   1
   RegLocation     =   "HKEY_CURRENT_USER\Software\Microsoft\Visual Basic\6.0"
   CmdLineSupport  =   -1  'True
End
Attribute VB_Name = "Connect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public FormDisplayed          As Boolean
Public VBInstance             As VBIDE.VBE
Dim mcbMenuCommandBar         As Office.CommandBarControl
Dim projectMng                As New ProjectManager
Public WithEvents MenuHandler As CommandBarEvents          'command bar event handler
Attribute MenuHandler.VB_VarHelpID = -1

Dim lnE As LanguageEditor
Dim cnE As ConfigurationEditor

Friend Function getp() As ProjectProperties
    Set getp = PojectProperties(VBInstance.ActiveVBProject)
End Function

Sub Hide()
    
    On Error Resume Next
    
    FormDisplayed = False
    Call Unload(projectMng)
    Call cnE.CloseForm
    Call lnE.CloseForm
   
End Sub

Sub Show()
  
    On Error Resume Next
    
    Set ProjectManager.VBInstance = VBInstance
    Set ProjectManager.Connect = Me
    FormDisplayed = True
    Call ProjectManager.Show
   
End Sub

Public Sub ShowConfigurationEditor()
    If cnE Is Nothing Then Set cnE = ConfigurationEditor
    Call cnE.Show
End Sub
Public Sub ShowLanguageEditor(Optional Path)
    If lnE Is Nothing Then Set lnE = LanguageEditor
    Call lnE.Show
End Sub
Public Sub HideConfigurationEditor()
    Call cnE.CloseForm
End Sub
Public Sub HideLanguageEditor()
    Call lnE.CloseForm
End Sub

'------------------------------------------------------
'this method adds the Add-In to VB
'------------------------------------------------------
Private Sub AddinInstance_OnConnection(ByVal Application As Object, ByVal ConnectMode As AddInDesignerObjects.ext_ConnectMode, ByVal AddInInst As Object, custom() As Variant)
    On Error GoTo error_handler
    
    'save the vb instance
    Set VBInstance = Application
    Set C = Me
    
    'this is a good place to set a breakpoint and
    'test various addin objects, properties and methods
    Debug.Print VBInstance.FullName


    If ConnectMode = ext_cm_External Then
        'Used by the wizard toolbar to start this wizard
        Me.Show
    Else
        Set mcbMenuCommandBar = AddToAddInCommandBar("MintAPI Project Manager")
        'sink the event

        Set Me.MenuHandler = VBInstance.Events.CommandBarEvents(mcbMenuCommandBar)
    End If

    If ConnectMode = ext_cm_AfterStartup Then
        If GetSetting(App.Title, "Settings", "DisplayOnConnect", "0") = "1" Then
            'set this to display the form on connect
            Me.Show
        End If
    End If
  
    Exit Sub
    
error_handler:
    
    MsgBox err.Description
    
End Sub

'------------------------------------------------------
'this method removes the Add-In from VB
'------------------------------------------------------
Private Sub AddinInstance_OnDisconnection(ByVal RemoveMode As AddInDesignerObjects.ext_DisconnectMode, custom() As Variant)
    On Error Resume Next
    
    'delete the command bar entry
    mcbMenuCommandBar.Delete
    
    'shut down the Add-In
    If FormDisplayed Then
        SaveSetting App.Title, "Settings", "DisplayOnConnect", "1"
        FormDisplayed = False
    Else
        SaveSetting App.Title, "Settings", "DisplayOnConnect", "0"
    End If
    
    Unload projectMng
    Set projectMng = Nothing

End Sub

Private Sub IDTExtensibility_OnStartupComplete(custom() As Variant)
    If GetSetting(App.Title, "Settings", "DisplayOnConnect", "0") = "1" Then
        'set this to display the form on connect
        Me.Show
    End If
End Sub

'this event fires when the menu is clicked in the IDE
Private Sub MenuHandler_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    Me.Show
End Sub

Function AddToAddInCommandBar(sCaption As String) As Office.CommandBarControl
    Dim cbMenuCommandBar As Office.CommandBarControl  'command bar object
    Dim cbMenu As Object
  
    On Error GoTo AddToAddInCommandBarErr
    
    'see if we can find the Add-Ins menu
    Set cbMenu = VBInstance.CommandBars("Add-Ins")
    If cbMenu Is Nothing Then
        'not available so we fail
        Exit Function
    End If
    
    'add it to the command bar
    Set cbMenuCommandBar = cbMenu.Controls.Add(1)
    'set the caption
    cbMenuCommandBar.Caption = sCaption
    
    Set AddToAddInCommandBar = cbMenuCommandBar
    
    Exit Function
    
AddToAddInCommandBarErr:

End Function

