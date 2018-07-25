VERSION 5.00
Begin {AC0714F6-3D04-11D1-AE7D-00A0C90F26F4} Connect 
   ClientHeight    =   9495
   ClientLeft      =   1740
   ClientTop       =   1545
   ClientWidth     =   11340
   _ExtentX        =   20003
   _ExtentY        =   16748
   _Version        =   393216
   Description     =   "MinAPI Intellicense for Visual Basic 6.0"
   DisplayName     =   "MinAPI Intellicense"
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

Public FormDisplayed                        As Boolean

Public Menu_Main As CommandBar
Public WithEvents Menu_MintAPIIntellicense As MenuItem
Attribute Menu_MintAPIIntellicense.VB_VarHelpID = -1
Public WithEvents Menu_MintAPIIntellicense_Tools As MenuItem
Attribute Menu_MintAPIIntellicense_Tools.VB_VarHelpID = -1
Public WithEvents Menu_MintAPIIntellicense_Tools_WidgetDesigner As MenuItem
Attribute Menu_MintAPIIntellicense_Tools_WidgetDesigner.VB_VarHelpID = -1
Public WithEvents Menu_MintAPIIntellicense_Tools_Globalization As MenuItem
Attribute Menu_MintAPIIntellicense_Tools_Globalization.VB_VarHelpID = -1
Public WithEvents Menu_MintAPIIntellicense_Settings As MenuItem
Attribute Menu_MintAPIIntellicense_Settings.VB_VarHelpID = -1
Public WithEvents Menu_MintAPIIntellicense_About As MenuItem
Attribute Menu_MintAPIIntellicense_About.VB_VarHelpID = -1


Public Sub InitializeEnvironment()
If GetSetting(App.Title, "Settings", "MenuJustInAddIns", "0") = "1" Then
    Set Menu_Main = VBInstanceWrapper.SelectCommandBarByName(MENU_ADDINS_NAME)
Else
    Set Menu_Main = VBInstanceWrapper.SelectCommandBarByName(MENU_NAME)
End If
    
    Dim Obj As CommandBarControl, Index As Long, Str As String
    For Each Obj In Menu_Main.Controls
        Index = Index + 1
        Str = Replace$(Obj.Caption, "&", "")
        If Str = MENU_TOOLS_NAME Then Exit For
        If Str = MENU_ADDINS_NAME Then Exit For
        If Str = MENU_WINDOW_NAME Then Exit For
        If Str = MENU_HELP_NAME Then Exit For
    Next
    
    '++++++++++++++++++++++++++
    '==========================
    Set Menu_MintAPIIntellicense = VBInstanceWrapper.CreateMenu(Menu_Main, MENU_ADDINS_MINTAPIINTERLLICENSE_CAPTION, msoControlPopup, Index)
    '==========================
    Set Menu_MintAPIIntellicense_Tools = MenuItem("&Tools", Menu_MintAPIIntellicense, True)
    '>>>
        Set Menu_MintAPIIntellicense_Tools_WidgetDesigner = MenuItem("&Widget Designer", Menu_MintAPIIntellicense_Tools)
        Set Menu_MintAPIIntellicense_Tools_Globalization = MenuItem("&Globalization", Menu_MintAPIIntellicense_Tools)
    '==========================
    Set Menu_MintAPIIntellicense_Settings = MenuItem("&Settings", Menu_MintAPIIntellicense)
    '==========================
    Set Menu_MintAPIIntellicense_About = MenuItem("&About MintIntellicense", Menu_MintAPIIntellicense)
    '==========================
    '++++++++++++++++++++++++++
End Sub
Public Sub FinilizeEnvironment()
    Call Menu_MintAPIIntellicense.Dispose
    Set Menu_MintAPIIntellicense = Nothing
    Set Connector = Nothing
End Sub


Public Sub Show()
    If Forms_Welcome Is Nothing Then _
        Set Forms_Welcome = New frmWelcome
    
    FormDisplayed = True
    Call Forms_Welcome.Show
End Sub
Public Sub Hide()
    FormDisplayed = False
    If Not Forms_Welcome Is Nothing Then _
        Call Forms_Welcome.Hide
End Sub

'------------------------------------------------------
'this method adds the Add-In to VB
'------------------------------------------------------
Private Sub AddinInstance_OnConnection(ByVal Application As Object, ByVal ConnectMode As AddInDesignerObjects.ext_ConnectMode, ByVal AddInInst As Object, custom() As Variant)
    On Error GoTo error_handler
    
    Set Connector = Me
    'save the vb instance
    Set VBInstanceWrapper = New VBInstanceWrapper
    Call VBInstanceWrapper.Constructor0(Application)
    
    If ConnectMode = ext_cm_External Then
        'Used by the wizard toolbar to start this wizard
        Call Me.Show
    Else
        Call InitializeEnvironment
    End If
  
    If ConnectMode = ext_cm_AfterStartup Then
        If GetSetting(App.Title, "Settings", "DisplayOnConnect", "0") = "1" Then
            'set this to display the form on connect
            Call Me.Show
        End If
    End If
    Exit Sub
error_handler:
    MsgBox Err.Description
End Sub

'------------------------------------------------------
'this method removes the Add-In from VB
'------------------------------------------------------
Private Sub AddinInstance_OnDisconnection(ByVal RemoveMode As AddInDesignerObjects.ext_DisconnectMode, custom() As Variant)
    On Error Resume Next
    
    'delete the command bar entry
    Call FinilizeEnvironment
    
    'shut down the Add-In
    If FormDisplayed Then
        SaveSetting App.Title, "Settings", "DisplayOnConnect", "1"
        FormDisplayed = False
    Else
        SaveSetting App.Title, "Settings", "DisplayOnConnect", "0"
    End If
End Sub

Private Sub IDTExtensibility_OnStartupComplete(custom() As Variant)
    If GetSetting(App.Title, "Settings", "DisplayOnConnect", "0") = "1" Then
        'set this to display the form on connect
        'Call Me.Show
    End If
End Sub

Private Sub Menu_MintAPIIntellicense_Tools_Globalization_Click(ByVal E As MintAPI.EventArgs)
    If Forms_Globalization Is Nothing Then _
        Set Forms_Globalization = New frmGlobalization
    Call Forms_Globalization.Show
End Sub
Private Sub Menu_MintAPIIntellicense_Tools_WidgetDesigner_Click(ByVal E As MintAPI.EventArgs)
    If Forms_WidgetDesigner Is Nothing Then _
        Set Forms_WidgetDesigner = New frmWidgetDesigner
    Call Forms_WidgetDesigner.Show
End Sub
Private Sub Menu_MintAPIIntellicense_Settings_Click(ByVal E As MintAPI.EventArgs)
    If Forms_Settings Is Nothing Then _
        Set Forms_Settings = New frmSettings
    Call Forms_Settings.Show(1)
End Sub
Private Sub Menu_MintAPIIntellicense_About_Click(ByVal E As MintAPI.EventArgs)
    If Forms_About Is Nothing Then _
        Set Forms_About = New frmAbout
    Call Forms_About.Show(1)
End Sub
