VERSION 5.00
Begin VB.Form Settings 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Settings"
   ClientHeight    =   9540
   ClientLeft      =   150
   ClientTop       =   135
   ClientWidth     =   13995
   BeginProperty Font 
      Name            =   "Consolas"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00404040&
   Icon            =   "Settings.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9540
   ScaleWidth      =   13995
   StartUpPosition =   1  'CenterOwner
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.Frame frm 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   8265
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11790
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   540
         Left            =   900
         TabIndex        =   22
         Top             =   6225
         Width           =   8265
         Begin VB.OptionButton rdShowAllAvailable 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Show All Available Tips Including New Application ,Projects ,Plugins And Everything Else"
            Height          =   240
            Left            =   0
            TabIndex        =   24
            Top             =   300
            Value           =   -1  'True
            Width           =   8265
         End
         Begin VB.OptionButton rdShowMiccyTips 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Only Show miccy Tips"
            Height          =   240
            Left            =   0
            TabIndex        =   23
            Top             =   0
            Width           =   2115
         End
      End
      Begin VB.CheckBox chkShowAppTips 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Show Application Tips"
         Height          =   240
         Left            =   600
         TabIndex        =   21
         Top             =   5925
         Value           =   1  'Checked
         Width           =   2190
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   540
         Left            =   900
         TabIndex        =   18
         Top             =   4950
         Width           =   4590
         Begin VB.OptionButton rdCheckAllPlgs 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Check For All Plugins Update (recomended)"
            Height          =   240
            Left            =   0
            TabIndex        =   20
            Top             =   0
            Width           =   3990
         End
         Begin VB.OptionButton rdOnlyPlugins 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Check Only miccy Plugins Update (more security)"
            Height          =   240
            Left            =   0
            TabIndex        =   19
            Top             =   300
            Value           =   -1  'True
            Width           =   4590
         End
      End
      Begin VB.CommandButton apply1Button 
         Caption         =   "&Apply"
         Height          =   390
         Left            =   10200
         TabIndex        =   17
         Top             =   675
         Width           =   1365
      End
      Begin VB.CommandButton save1Button 
         Caption         =   "&Save"
         Height          =   390
         Left            =   10200
         TabIndex        =   16
         Top             =   1125
         Width           =   1365
      End
      Begin VB.CheckBox chkAutoCheckUpdates 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Auto Check For miccy And Plugins Update (recomended)"
         Height          =   240
         Left            =   600
         TabIndex        =   15
         Top             =   4650
         Value           =   1  'Checked
         Width           =   4965
      End
      Begin VB.CheckBox chkAllowToConnectToNetwork 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Allow Plugins To Connect To Network Using miccy API (this not take effect on internal plugin Connections to the network)"
         Height          =   465
         Left            =   600
         TabIndex        =   14
         Top             =   4050
         Value           =   1  'Checked
         Width           =   8565
      End
      Begin VB.CheckBox chkAllowTerminateProccess 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Allow Plugins To Terminate The Proccess (recomended)"
         Height          =   240
         Left            =   600
         TabIndex        =   13
         Top             =   3675
         Value           =   1  'Checked
         Width           =   4965
      End
      Begin VB.CheckBox chkAllowInstallUninstallOthers 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Allow Plugins To Install/Uninstall Other Plugins"
         Height          =   240
         Left            =   600
         TabIndex        =   12
         Top             =   3300
         Width           =   4665
      End
      Begin VB.CheckBox chkAllowChangeGlobalSettings 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Allow Plugins To Take Changes In Global Settigns (weak privacy/security)"
         Height          =   240
         Left            =   600
         TabIndex        =   11
         Top             =   2925
         Width           =   6765
      End
      Begin VB.CheckBox chkLoadPluginsOnStartup 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Load All Plugins On Startup To Collect Plugins Information (not recomended)"
         Height          =   240
         Left            =   600
         TabIndex        =   10
         Top             =   2250
         Width           =   7065
      End
      Begin VB.CommandButton saveButton 
         Caption         =   "&Save"
         Height          =   390
         Left            =   10125
         TabIndex        =   9
         Top             =   7200
         Width           =   1365
      End
      Begin VB.CommandButton applyButton 
         Caption         =   "&Apply"
         Height          =   390
         Left            =   10125
         TabIndex        =   8
         Top             =   6750
         Width           =   1365
      End
      Begin VB.CheckBox chkmultiInstance 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Run Multi Instanced"
         Height          =   240
         Left            =   600
         TabIndex        =   7
         Top             =   1875
         Width           =   2040
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Restart This Option On New Instance (Recomended)"
         Height          =   240
         Left            =   600
         TabIndex        =   6
         Top             =   1200
         Value           =   1  'Checked
         Width           =   4665
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Buffer Data On Memory"
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   2625
         TabIndex        =   3
         Top             =   900
         Width           =   2190
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Buffer Data On Disk"
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   2625
         TabIndex        =   2
         Top             =   675
         Value           =   -1  'True
         Width           =   2040
      End
      Begin VB.Line grid 
         Index           =   2
         Visible         =   0   'False
         X1              =   850
         X2              =   850
         Y1              =   75
         Y2              =   10225
      End
      Begin VB.Label lbls 
         BackStyle       =   0  'Transparent
         Caption         =   "Donate us to make this application(s) better ,faster and more usefull ,and you can be a member just Typing login.deftro.com"
         ForeColor       =   &H00808080&
         Height          =   450
         Index           =   3
         Left            =   850
         TabIndex        =   25
         Top             =   7275
         Width           =   9045
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FBE3C4&
         Index           =   2
         X1              =   525
         X2              =   10125
         Y1              =   5700
         Y2              =   5700
      End
      Begin VB.Line grid 
         Index           =   1
         Visible         =   0   'False
         X1              =   10050
         X2              =   10050
         Y1              =   0
         Y2              =   10150
      End
      Begin VB.Line grid 
         Index           =   0
         Visible         =   0   'False
         X1              =   600
         X2              =   600
         Y1              =   0
         Y2              =   10150
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FBE3C4&
         Index           =   1
         X1              =   525
         X2              =   10125
         Y1              =   2700
         Y2              =   2700
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FBE3C4&
         Index           =   0
         X1              =   525
         X2              =   10125
         Y1              =   1650
         Y2              =   1650
      End
      Begin VB.Label lbls 
         BackStyle       =   0  'Transparent
         Caption         =   $"Settings.frx":57E2
         ForeColor       =   &H00808080&
         Height          =   825
         Index           =   2
         Left            =   5700
         TabIndex        =   5
         Top             =   675
         Width           =   4395
      End
      Begin VB.Label lbls 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tip:"
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   0
         Left            =   5325
         TabIndex        =   4
         Top             =   675
         Width           =   360
      End
      Begin VB.Label lbls 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Data Buffer Location:"
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   1
         Left            =   600
         TabIndex        =   1
         Top             =   675
         Width           =   1890
      End
   End
   Begin VB.Menu m_menu 
      Caption         =   "&Menu"
      Visible         =   0   'False
      Begin VB.Menu m_menu_clipboard 
         Caption         =   "&Clipboard Operations"
         Begin VB.Menu m_menu_clip_copy 
            Caption         =   "&Copy"
            Shortcut        =   ^C
         End
         Begin VB.Menu m_menu_clip_cut 
            Caption         =   "C&ut"
            Shortcut        =   ^X
         End
         Begin VB.Menu m_menu_clip_paste 
            Caption         =   "&Paste"
            Shortcut        =   ^V
         End
         Begin VB.Menu m_menu_clip_clear 
            Caption         =   "C&lear Clipboard"
         End
      End
      Begin VB.Menu m_menu_io 
         Caption         =   "&IO Operations"
         Begin VB.Menu m_menu_io_new 
            Caption         =   "&New"
            Shortcut        =   ^N
         End
         Begin VB.Menu m_menu_io_open 
            Caption         =   "&Open"
            Shortcut        =   ^O
         End
         Begin VB.Menu m_menu_io_save 
            Caption         =   "&Save"
            Shortcut        =   ^S
         End
         Begin VB.Menu m_menu_io_saveas 
            Caption         =   "Save &As"
         End
         Begin VB.Menu m_menu_io_close 
            Caption         =   "&Close"
            Shortcut        =   ^Q
         End
      End
      Begin VB.Menu m_menu_edit 
         Caption         =   "&Edit Operations"
         Begin VB.Menu m_menu_edit_undo 
            Caption         =   "&Undo"
            Shortcut        =   ^Z
         End
         Begin VB.Menu m_menu_edit_redo 
            Caption         =   "&Redo"
            Shortcut        =   ^Y
         End
         Begin VB.Menu m_menu_edti_selectall 
            Caption         =   "&Select All"
            Shortcut        =   ^A
         End
         Begin VB.Menu m_menu_edit_history 
            Caption         =   "View &History"
            Shortcut        =   ^H
         End
         Begin VB.Menu m_menu_edti_clear 
            Caption         =   "&Clear History"
         End
      End
      Begin VB.Menu m_menu_sep1 
         Caption         =   "-"
      End
      Begin VB.Menu m_menu_closeall 
         Caption         =   "&Close All Open Plugins"
      End
      Begin VB.Menu m_menu_actions 
         Caption         =   "Action"
         Index           =   0
         Visible         =   0   'False
      End
   End
   Begin VB.Menu m_actions 
      Caption         =   "&Actions"
      Visible         =   0   'False
      Begin VB.Menu m_actions_search 
         Caption         =   "Search Online For &Plugins"
      End
      Begin VB.Menu m_actions_actions 
         Caption         =   "Action"
         Index           =   0
         Visible         =   0   'False
      End
   End
   Begin VB.Menu m_tools 
      Caption         =   "&Tools"
      Visible         =   0   'False
      Begin VB.Menu m_tools_install 
         Caption         =   "&Install New Plugins"
      End
      Begin VB.Menu m_tools_uninstall 
         Caption         =   "&Uninstall Plugins"
      End
      Begin VB.Menu m_tools_sep1 
         Caption         =   "-"
      End
      Begin VB.Menu m_tools_manage 
         Caption         =   "&Plugin Manager"
         Shortcut        =   ^B
      End
   End
   Begin VB.Menu m_config 
      Caption         =   "&Configuration"
      Visible         =   0   'False
      Begin VB.Menu m_config_settings 
         Caption         =   "&Settings"
         Shortcut        =   ^G
      End
   End
   Begin VB.Menu m_help 
      Caption         =   "&Help"
      Visible         =   0   'False
      Begin VB.Menu m_help_docs 
         Caption         =   "&Documentation"
         Shortcut        =   ^{F4}
      End
      Begin VB.Menu m_help_search 
         Caption         =   "&Search"
      End
      Begin VB.Menu m_help_sep1 
         Caption         =   "-"
      End
      Begin VB.Menu m_help_online 
         Caption         =   "&Online Documentation"
         Shortcut        =   ^{F11}
      End
      Begin VB.Menu m_help_support 
         Caption         =   "S&upport"
      End
      Begin VB.Menu m_help_forum 
         Caption         =   "&Forum"
      End
      Begin VB.Menu m_help_myAccount 
         Caption         =   "&My Account"
         Shortcut        =   ^{F3}
      End
      Begin VB.Menu m_help_sep2 
         Caption         =   "-"
      End
      Begin VB.Menu m_help_donate 
         Caption         =   "&Donate (even 1$ makes us happy)"
      End
      Begin VB.Menu m_help_reg 
         Caption         =   "&Registration (free)"
         Visible         =   0   'False
      End
      Begin VB.Menu m_help_about 
         Caption         =   "&About miccy Ultimate Tools"
         Shortcut        =   {F1}
      End
   End
End
Attribute VB_Name = "Settings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub m_config_settings_Click(): Call mForm.ScrollThis(frm): End Sub
Private Sub m_help_about_Click(): Call About.Show(1): End Sub
Private Sub m_help_docs_Click()
    '
End Sub
Private Sub m_help_forum_Click()
    '
End Sub
Private Sub m_help_myAccount_Click()
    '
End Sub
Private Sub m_help_online_Click()
    '
End Sub
Private Sub m_help_search_Click()
    '
End Sub
Private Sub m_help_support_Click()
    '
End Sub

Private Sub writeSettingsToGP()
    gp.MultiInstance = (chkmultiInstance.Value = vbChecked)
    gp.LoadAllPluginsInStartup = (chkLoadPluginsOnStartup.Value = vbChecked)
    gp.AllowPluginsToChangeGlobalSettings = (chkAllowChangeGlobalSettings.Value = vbChecked)
    gp.AllowPluginsToInstallOrUninstallOthers = (chkAllowInstallUninstallOthers.Value = vbChecked)
    gp.AllowPluginsToTerminateProccess = (chkAllowTerminateProccess.Value = vbChecked)
    gp.AllowPluginsToConnectToNetwork = (chkAllowToConnectToNetwork.Value = vbChecked)
    If (chkAutoCheckUpdates.Value = vbChecked) Then
        If rdCheckAllPlgs.Value Then
            gp.AllowPluginsToTerminateProccess = &H1
        Else
            gp.AllowPluginsToTerminateProccess = &H10
        End If
    End If
    If (chkShowAppTips.Value = vbChecked) Then
        If rdShowMiccyTips.Value Then
            gp.AllowPluginsToTerminateProccess = &H1
        Else
            gp.AllowPluginsToTerminateProccess = &H10
        End If
    End If
    
End Sub

Private Sub applyButton_Click()
    Call writeSettingsToGP
End Sub
Private Sub saveButton_Click()
    Call writeSettingsToGP
    gpMustSave = True
    Call SaveConfig
End Sub

Private Sub save1Button_Click(): Call saveButton_Click: End Sub
Private Sub apply1Button_Click(): Call applyButton_Click: End Sub

Private Sub chkAutoCheckUpdates_Click()
    rdCheckAllPlgs.Enabled = (chkAutoCheckUpdates.Value = vbChecked)
    rdOnlyPlugins.Enabled = (chkAutoCheckUpdates.Value = vbChecked)
End Sub
Private Sub chkShowAppTips_Click()
    rdShowAllAvailable.Enabled = (chkShowAppTips.Value = vbChecked)
    rdShowMiccyTips.Enabled = (chkShowAppTips.Value = vbChecked)
End Sub
