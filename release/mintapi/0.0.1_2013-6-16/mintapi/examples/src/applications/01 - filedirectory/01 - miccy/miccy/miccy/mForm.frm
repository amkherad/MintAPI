VERSION 5.00
Begin VB.Form mForm 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "miccy Ultimate Tools [Advanced]"
   ClientHeight    =   8490
   ClientLeft      =   -45
   ClientTop       =   -375
   ClientWidth     =   12600
   ControlBox      =   0   'False
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
   Icon            =   "mForm.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   Picture         =   "mForm.frx":57E2
   ScaleHeight     =   566
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   840
   StartUpPosition =   2  'CenterScreen
   WhatsThisHelp   =   -1  'True
   Begin VB.Timer scrollTimer 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   7650
      Top             =   150
   End
   Begin VB.ComboBox sorter 
      Height          =   315
      ItemData        =   "mForm.frx":E594
      Left            =   7575
      List            =   "mForm.frx":E5BC
      Style           =   2  'Dropdown List
      TabIndex        =   16
      Top             =   885
      Width           =   1665
   End
   Begin VB.TextBox searchText 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   270
      Left            =   9450
      TabIndex        =   14
      Text            =   "Search Plugins Here"
      Top             =   900
      Width           =   2370
   End
   Begin VB.Timer mnuTimer 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   7125
      Top             =   150
   End
   Begin VB.Frame frm 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7050
      Left            =   195
      TabIndex        =   2
      Top             =   1230
      Width           =   11850
      Begin VB.Frame plgWinfrm 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1065
         Left            =   0
         TabIndex        =   19
         Top             =   3600
         Visible         =   0   'False
         Width           =   11865
      End
      Begin VB.Frame plgfrm 
         Appearance      =   0  'Flat
         BackColor       =   &H00FBD2AE&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1065
         Left            =   0
         TabIndex        =   15
         Top             =   2475
         Visible         =   0   'False
         Width           =   11865
      End
      Begin VB.Frame infrm 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2445
         Left            =   0
         TabIndex        =   12
         Top             =   0
         Width           =   11865
         Begin miccyUltimateTools.Plugin p 
            Height          =   1665
            Index           =   0
            Left            =   0
            TabIndex        =   18
            Top             =   0
            Width           =   11925
            _ExtentX        =   21034
            _ExtentY        =   2937
         End
         Begin VB.Label noplugin 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "No Plugin Installed"
            BeginProperty Font 
               Name            =   "Consolas"
               Size            =   20.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   480
            Left            =   0
            TabIndex        =   13
            Top             =   3075
            Width           =   11805
         End
      End
   End
   Begin VB.Label icon 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   840
      Left            =   225
      TabIndex        =   0
      Top             =   0
      Width           =   690
   End
   Begin VB.Image mastersettingsButton 
      Height          =   240
      Left            =   6000
      Picture         =   "mForm.frx":E636
      ToolTipText     =   "Master Settings"
      Top             =   900
      Width           =   240
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00E0E0E0&
      X1              =   420
      X2              =   420
      Y1              =   60
      Y2              =   78
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00E0E0E0&
      X1              =   395
      X2              =   395
      Y1              =   60
      Y2              =   78
   End
   Begin VB.Image infoButton 
      Height          =   240
      Left            =   6375
      Picture         =   "mForm.frx":E8B6
      ToolTipText     =   "Show Information About Selected Plugin"
      Top             =   900
      Width           =   240
   End
   Begin VB.Image updateButton 
      Height          =   240
      Left            =   6975
      Picture         =   "mForm.frx":ECD6
      ToolTipText     =   "Update Selected Plugin"
      Top             =   900
      Width           =   240
   End
   Begin VB.Image uninstallButton 
      Height          =   240
      Left            =   6675
      Picture         =   "mForm.frx":EF2C
      ToolTipText     =   "Uninstall Selected Plugin"
      Top             =   900
      Width           =   240
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00E0E0E0&
      X1              =   370
      X2              =   370
      Y1              =   60
      Y2              =   78
   End
   Begin VB.Image scrMid 
      Height          =   105
      Left            =   12180
      Picture         =   "mForm.frx":F14D
      Top             =   4425
      Width           =   135
   End
   Begin VB.Line scrollDown 
      BorderColor     =   &H0000C000&
      Visible         =   0   'False
      X1              =   808
      X2              =   826
      Y1              =   552
      Y2              =   552
   End
   Begin VB.Line scrollTop 
      BorderColor     =   &H0000C000&
      Visible         =   0   'False
      X1              =   808
      X2              =   826
      Y1              =   55
      Y2              =   55
   End
   Begin VB.Image scr 
      Height          =   7110
      Left            =   12180
      Picture         =   "mForm.frx":F2FD
      Stretch         =   -1  'True
      Top             =   960
      Width           =   135
   End
   Begin VB.Image scrDown 
      Height          =   60
      Left            =   12180
      Picture         =   "mForm.frx":F44C
      Top             =   8145
      Width           =   135
   End
   Begin VB.Image scrTop 
      Height          =   60
      Left            =   12180
      Picture         =   "mForm.frx":F5BB
      Top             =   900
      Width           =   135
   End
   Begin VB.Image openPlugins 
      Height          =   240
      Left            =   5625
      Picture         =   "mForm.frx":F72A
      ToolTipText     =   "List All Open Plugins"
      Top             =   900
      Width           =   240
   End
   Begin VB.Image backButton 
      Height          =   240
      Index           =   2
      Left            =   4800
      Picture         =   "mForm.frx":F9A4
      ToolTipText     =   "<Previeus Page"
      Top             =   450
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image nextButton 
      Height          =   240
      Index           =   2
      Left            =   5025
      Picture         =   "mForm.frx":FA33
      ToolTipText     =   "Next Page>"
      Top             =   450
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image backButton 
      Height          =   240
      Index           =   1
      Left            =   4800
      Picture         =   "mForm.frx":FAC2
      ToolTipText     =   "<Previeus Page"
      Top             =   600
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image nextButton 
      Height          =   240
      Index           =   1
      Left            =   5025
      Picture         =   "mForm.frx":FB51
      ToolTipText     =   "Next Page>"
      Top             =   600
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image homeButton 
      Height          =   240
      Left            =   5250
      Picture         =   "mForm.frx":FBE0
      ToolTipText     =   "Plugin Picker Page"
      Top             =   900
      Width           =   240
   End
   Begin VB.Image settingsButton 
      Height          =   240
      Left            =   7275
      Picture         =   "mForm.frx":FC8B
      ToolTipText     =   "Configure Selected Plugin"
      Top             =   900
      Width           =   240
   End
   Begin VB.Image nextButton 
      Height          =   240
      Index           =   0
      Left            =   5025
      Picture         =   "mForm.frx":FEDA
      ToolTipText     =   "Next Page>"
      Top             =   900
      Width           =   240
   End
   Begin VB.Image backButton 
      Height          =   240
      Index           =   0
      Left            =   4800
      Picture         =   "mForm.frx":FF69
      ToolTipText     =   "<Previeus Page"
      Top             =   900
      Width           =   240
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   285
      Left            =   9375
      Top             =   900
      Width           =   240
   End
   Begin VB.Image searchGo 
      Height          =   270
      Left            =   11850
      Picture         =   "mForm.frx":FFF8
      Stretch         =   -1  'True
      Top             =   900
      Width           =   240
   End
   Begin VB.Image search 
      Height          =   420
      Left            =   9375
      Picture         =   "mForm.frx":100D7
      Stretch         =   -1  'True
      Top             =   825
      Width           =   2640
   End
   Begin VB.Image searchRight 
      Height          =   420
      Left            =   12000
      Picture         =   "mForm.frx":10243
      Top             =   825
      Width           =   90
   End
   Begin VB.Image searchLeft 
      Height          =   420
      Left            =   9300
      Picture         =   "mForm.frx":1045B
      Top             =   825
      Width           =   90
   End
   Begin VB.Label mnu 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Help"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   4
      Left            =   3300
      TabIndex        =   9
      Top             =   900
      Width           =   750
   End
   Begin VB.Label mnu 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Configuration"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   3
      Left            =   2175
      TabIndex        =   8
      Top             =   900
      Width           =   1350
   End
   Begin VB.Label mnu 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Tools"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   2
      Left            =   1575
      TabIndex        =   7
      Top             =   900
      Width           =   750
   End
   Begin VB.Label mnu 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Actions"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   1
      Left            =   975
      TabIndex        =   6
      Top             =   915
      Width           =   750
   End
   Begin VB.Label mnu 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Menu"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   0
      Left            =   450
      TabIndex        =   5
      Top             =   915
      Width           =   600
   End
   Begin VB.Image menuhRight 
      Height          =   420
      Left            =   4350
      Picture         =   "mForm.frx":1066C
      Top             =   825
      Width           =   180
   End
   Begin VB.Image menuhLeft 
      Height          =   420
      Left            =   375
      Picture         =   "mForm.frx":10929
      Top             =   825
      Width           =   180
   End
   Begin VB.Image menuh 
      Height          =   420
      Left            =   525
      Picture         =   "mForm.frx":10BCF
      Stretch         =   -1  'True
      Top             =   825
      Width           =   3840
   End
   Begin VB.Image menu 
      Height          =   420
      Left            =   300
      Picture         =   "mForm.frx":10D67
      Stretch         =   -1  'True
      Top             =   825
      Width           =   4365
   End
   Begin VB.Image menuRight 
      Height          =   420
      Left            =   4650
      Picture         =   "mForm.frx":10ED3
      Top             =   825
      Width           =   90
   End
   Begin VB.Image menuLeft 
      Height          =   420
      Left            =   225
      Picture         =   "mForm.frx":110EB
      Top             =   825
      Width           =   90
   End
   Begin VB.Image closeBTN 
      Height          =   240
      Left            =   12150
      Picture         =   "mForm.frx":112FC
      Top             =   150
      Width           =   240
   End
   Begin VB.Image minimizeBTN 
      Height          =   240
      Left            =   11925
      Picture         =   "mForm.frx":136CE
      Top             =   150
      Width           =   240
   End
   Begin VB.Label scrRectEventer 
      BackStyle       =   0  'Transparent
      Height          =   7485
      Left            =   12150
      TabIndex        =   17
      Top             =   825
      Width           =   270
   End
   Begin VB.Shape scrollRect 
      BackColor       =   &H0000C000&
      BorderColor     =   &H0000C000&
      FillColor       =   &H0000C000&
      FillStyle       =   0  'Solid
      Height          =   7185
      Left            =   12120
      Top             =   975
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Label title 
      BackStyle       =   0  'Transparent
      Height          =   840
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   12615
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright by deftro.com"
      ForeColor       =   &H00FBD2AE&
      Height          =   240
      Left            =   10275
      TabIndex        =   11
      Top             =   405
      Width           =   2115
   End
   Begin VB.Label donater 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Donate Us To Make Applications Even Better"
      ForeColor       =   &H00FBD2AE&
      Height          =   240
      Left            =   8550
      TabIndex        =   10
      Top             =   600
      Width           =   3840
   End
   Begin VB.Label titleCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "miccy Ultimate Tools [Advanced]"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   330
      Left            =   1125
      TabIndex        =   3
      Top             =   150
      Width           =   4650
   End
   Begin VB.Label versionCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "version 1.0.0.2012"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   2475
      TabIndex        =   4
      Top             =   450
      Width           =   1620
   End
End
Attribute VB_Name = "mForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------
'--################################################################################--
'--################################################################################--
'------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------
'global definitions

Option Explicit
Const FORMWIDTH As Long = 840
Const FORMHEIGHT As Long = 566

Const MNUEFFECTTIMERINTERVAL As Long = 10
Const MNUEFFECTLATENCY As Long = 200 'means 2 seconds.
Const SCREFFECTTIMERINTERVAL As Long = 10
Const SCREFFECTLATENCY As Long = 100 'means 2 seconds.

Dim fX As Long, fY As Long

Dim mnuEfcLastTargetID As Long
Dim mnuEfcTargetID As Long
Dim mnuEfcLeft As Double
Dim mnuEfcLeftIncrease As Double
Dim mnuEfcWidth As Double
Dim mnuEfcWidthIncrease As Double
Const mnuPADDINGWIDTH As Long = 14 'must integerally divide by 2
Const mnuPADDINGWIDTH_2 As Long = mnuPADDINGWIDTH / 2

Const DEFAULTSEARCHTEXT As String = "Search Plugins Here"

Const MINPIXMOVTOSCROLL As Long = 5
Const SCROLLMININC As Long = 5
Dim scrMax As Long, scrValue As Long
Dim scrEfcTop As Double, scrConTop As Double, scrConIncrease As Double
Dim scrEfcTopIncrease As Double
Dim scrVisible As Boolean
Dim WithEvents scrollerObject As Frame
Attribute scrollerObject.VB_VarHelpID = -1
Dim LastscrollerObject As Frame
Attribute LastscrollerObject.VB_VarHelpID = -1
Dim firstMY As Long, secondMY As Long
Dim realScrollHeight As Long

Dim candirectFrame As Boolean
Dim pcandirectFrame As Boolean

Dim history() As String
Dim historyCount As Long
Dim historyMax As Long

Dim showingPlugins As Long

'------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------
'--################################################################################--
'--################################################################################--
'------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------
'main form

Private Sub Form_Load()
    mnuEfcLeftIncrease = 0: mnuEfcWidthIncrease = 0
    Call Me.Move(Me.Left, Me.top, FORMWIDTH * 15, FORMHEIGHT * 15)
    Dim hRGN As Long
    hRGN = API_CreateRoundRectRgn(0, 0, FORMWIDTH, FORMHEIGHT, 16, 16)
    Call API_SetWindowRgn(hwnd, hRGN, True)
    Call API_DeleteObject(hRGN)
    
    mnuTimer.Interval = MNUEFFECTTIMERINTERVAL
    scrollTimer.Interval = SCREFFECTTIMERINTERVAL
    
    sorter.ListIndex = 0
    
    frm.Width = scrollRect.Left - frm.Left
    
    pcandirectFrame = True
    candirectFrame = True
    
    Call API_SetParent(Settings.frm.hwnd, frm.hwnd)
    Settings.frm.Visible = False
    Call Settings.frm.Move(0, 0, frm.Width * 15)
    
    historyMax = gp.CountHistoryItems
    If historyMax < 5 Then historyMax = 5
    
    Call InitializeMenu
    Call InitializePlugins
    Call InitializeScroller
    Call InitializeEnvironment
End Sub

Private Sub title_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single): fX = X: fY = Y: End Sub
Private Sub title_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        Call Me.Move(Me.Left + X - fX, Me.top + Y - fY)
    End If
    Call mnuGoto(0)
End Sub
Private Sub closeBTN_Click(): Call Unload(Me): End Sub
Private Sub icon_DblClick(): Call closeBTN_Click: End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer): Call EndApp: End Sub

'Private Sub infrm_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single): Call mnuGoto(0): End Sub
'Private Sub frm_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single): Call mnuGoto(0): End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single): Call mnuGoto(0): End Sub

Private Sub minimizeBTN_Click(): Me.WindowState = 1: End Sub

Private Sub InitializePlugins()
    Debug.Print "Initializing Plugins"
    Call InitializePluginGUIItems
    Debug.Print "Plugins Initialized"
End Sub
Private Sub InitializePluginGUIItems()
    Debug.Print "Initializing Plugins GUI Items"
    plgfrm.top = frm.top
    infrm.top = frm.top
    plgfrm.Width = frm.Width * 15
    plgfrm.Height = frm.Height * 15
    infrm.Width = plgfrm.Width
    
    p(0).Width = infrm.Width
    noplugin.Width = infrm.Width
    noplugin.top = ((frm.Height * 15) / 2) - (noplugin.Height / 2)
    
    Dim i As Long, maxHeight As Long
    maxHeight = p(0).Height
    showingPlugins = 40
    For i = 1 To showingPlugins - 1
        Call Load(p(i))
        p(i).top = p(i - 1).top + p(i - 1).Height
        p(i).Visible = True
        maxHeight = maxHeight + p(i).Height '+ 1500
        p(i).Caption = "Plugin Name " & i
        p(i).Version = i
        p(i).Company = "Company " & i
    Next
    infrm.Height = maxHeight
    Debug.Print "Plugins GUI Items Initialized"
End Sub
Private Sub InitializeMenu()
    Debug.Print "Initializing Menu"
    Dim i As Long, initLeft As Long
    initLeft = mnu(0).Left
    For i = mnu.LBound To mnu.UBound
        Call mnu(i).Move(initLeft, 62)
        mnu(i).Width = Me.TextWidth(mnu(i).Caption) + mnuPADDINGWIDTH
        initLeft = initLeft + mnu(i).Width
    Next
    menuh.Left = mnu(0).Left + (mnuPADDINGWIDTH / 2)
    menuh.Width = mnu(0).Width - mnuPADDINGWIDTH
    menuhLeft.Left = menuh.Left - menuhLeft.Width
    menuhRight.Left = menuh.Left + menuh.Width
    Debug.Print "Menu Initialized"
End Sub
Private Sub InitializeScroller()
    Debug.Print "Initializing Scroller"
    Call ScrollThis(infrm) 'Settings.frm
    Debug.Print "Scroller Initialized"
End Sub
Private Sub InitializeEnvironment()
    Debug.Print "Initializing Environment"
    searchText.Text = DEFAULTSEARCHTEXT
    searchText.Visible = True
    Call SetBackEnabled(False)
    Call SetNextEnabled(False)
    '==========================================
    Call Show
    Call searchText.SetFocus
    Debug.Print "Environment Initialized"
End Sub

'------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------
'--################################################################################--
'--################################################################################--
'------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------
'menu effect area

Private Sub mnu_Click(Index As Integer)
    Const EXTRAMENUTOP As Long = 6
    Dim mLeft As Long, mTop As Long
    mLeft = Me.ScaleLeft + mnu(Index).Left
    mTop = Me.ScaleTop + mnu(Index).top + mnu(Index).Height + EXTRAMENUTOP
    Select Case Index
        Case 0
            Call PopupMenu(Settings.m_menu, , mLeft, mTop)
        Case 1
            Call PopupMenu(Settings.m_actions, , mLeft, mTop)
        Case 2
            Call PopupMenu(Settings.m_tools, , mLeft, mTop)
        Case 3
            Call PopupMenu(Settings.m_config, , mLeft, mTop)
        Case 4
            Call PopupMenu(Settings.m_help, , mLeft, mTop)
    End Select
End Sub

Private Sub mnu_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call mnuGoto(Index)
End Sub
Private Sub mnuGoto(ByVal Index As Long)
    If Index = mnuEfcTargetID Then Exit Sub
    mnuEfcLastTargetID = mnuEfcTargetID
    mnuEfcTargetID = Index
    mnuEfcLeft = menuh.Left '+ mnuPADDINGWIDTH_2
    mnuEfcWidth = menuh.Width '- mnuPADDINGWIDTH
    Const DINGS As Double = MNUEFFECTLATENCY / MNUEFFECTTIMERINTERVAL
    mnuEfcLeftIncrease = ((mnu(mnuEfcTargetID).Left + mnuPADDINGWIDTH_2) - mnuEfcLeft) / DINGS
    mnuEfcWidthIncrease = ((mnu(mnuEfcTargetID).Width - mnuPADDINGWIDTH) - mnuEfcWidth) / DINGS
    If Not (CLng(mnuEfcLeft) = mnu(mnuEfcTargetID).Left) Then
        mnuTimer.Enabled = True
    End If
End Sub
Private Sub mnuTimer_Timer()
    Dim lastLeft As Double, lastWidth As Double, bufLeft As Long
    lastLeft = mnuEfcLeft: lastWidth = mnuEfcWidth
    mnuEfcLeft = mnuEfcLeft + mnuEfcLeftIncrease
    mnuEfcWidth = mnuEfcWidth + mnuEfcWidthIncrease
    bufLeft = mnu(mnuEfcTargetID).Left + mnuPADDINGWIDTH_2
    If lastLeft >= bufLeft Then
        If CLng(mnuEfcLeft) <= bufLeft Then
            GoTo StopIt
        End If
    End If
    If lastLeft <= bufLeft Then
        If CLng(mnuEfcLeft) >= bufLeft Then
            GoTo StopIt
        End If
    End If
    GoTo NoStopIt
StopIt:
    mnuEfcLeft = bufLeft
    mnuEfcWidth = mnu(mnuEfcTargetID).Width - mnuPADDINGWIDTH
    mnuTimer.Enabled = False
NoStopIt:
    menuh.Left = CLng(mnuEfcLeft) ' + mnuPADDINGWIDTH_2
    menuh.Width = CLng(mnuEfcWidth) '- mnuPADDINGWIDTH
    menuhLeft.Left = menuh.Left - menuhLeft.Width
    menuhRight.Left = menuh.Left + menuh.Width
End Sub

'------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------
'--################################################################################--
'--################################################################################--
'------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------
'scroll area

Private Property Get ScrollVisible() As Boolean
    ScrollVisible = scrVisible
End Property
Private Property Let ScrollVisible(Value As Boolean)
    scrVisible = Value
    scr.Visible = Value
    scrTop.Visible = Value
    scrDown.Visible = Value
    scrMid.Visible = Value
    ScrollValue = 0
End Property

Public Sub ScrollThis(ByVal Obj As Object)
    Debug.Print "Executing Function ScrollThis()"
    '----------------------------------------
    If Not pcandirectFrame Then Exit Sub
    If scrollerObject Is Obj Then Exit Sub
    If Not scrollerObject Is Nothing Then scrollerObject.Visible = False
    Set LastscrollerObject = scrollerObject
    Set scrollerObject = Obj
    '----------------------------------------
    If Not scrollerObject Is Nothing Then Call appendHistory(scrollerObject)
    
    scrollerObject.Visible = True
    scrollerObject.top = 0
    Dim maxim As Long
    maxim = (Obj.Height / 15) - frm.Height
    If maxim <= 0 Then
        ScrollVisible = False
    Else
        ScrollMax = maxim
        If Not ScrollVisible Then ScrollVisible = True
        ScrollValue = 0
    End If
    Call scrollerObject.ZOrder(ZOrderConstants.vbBringToFront)
    Debug.Print "Execution Of Function ScrollThis() Had Ended"
End Sub
Private Property Get ScrollMax() As Long
    ScrollMax = scrMax
End Property
Private Property Let ScrollMax(Value As Long)
    If Value <= 0 Then Value = 1
    If Value >= 10000000 Then Value = 10000000
    scrMax = Value
    Dim bufHeight As Long
    bufHeight = scrollRect.Height - (Value / SCROLLMININC)
    realScrollHeight = bufHeight
    If bufHeight < scrMid.Height + 2 Then bufHeight = scrMid.Height + 2
    scr.Height = bufHeight
    If scrValue > scrMax Then scrValue = scrMax
    'If Not ScrollVisible Then Exit Property
End Property
Private Property Get ScrollValue() As Long
    ScrollValue = scrValue
End Property
Private Property Let ScrollValue(Value As Long)
    If Value < 0 Then Value = 0
    If Value > scrMax Then Value = scrMax
    scrValue = Value
    If Not ScrollVisible Then Exit Property
    If Not scrollerObject Is Nothing Then scrollerObject.top = -(scrValue * 15)
    scr.top = scrollRect.top + (((scrollRect.Height - scr.Height) * scrValue) / scrMax) ' + scr.Height / 2
    scrMid.top = scr.top + ((scr.Height / 2) - (scrMid.Height / 2))
    scrTop.top = scr.top - scrTop.Height
    scrDown.top = scr.top + scr.Height
End Property
Private Sub scrollTimer_Timer()
    scrConTop = scrConTop + scrConIncrease
    ScrollValue = scrConTop
    If ScrollValue = 0 Or ScrollValue = scrollerObject.Height Then scrollTimer.Enabled = False
End Sub

Private Sub scrollerObject_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        firstMY = Y / 15
    End If
End Sub
Private Sub scrollerObject_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        secondMY = Y / 15
        scrollTimer.Enabled = False
        ScrollValue = ScrollValue - (secondMY - firstMY)
    End If
End Sub
Private Sub scrollerObject_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
'        If (secondMY - firstMY) > MINPIXMOVTOSCROLL Then
'            scrConIncrease = -(CDbl(secondMY - firstMY))
'            scrollTimer.Enabled = True
'        ElseIf (firstMY - secondMY) > MINPIXMOVTOSCROLL Then
'            scrConIncrease = CDbl(firstMY - secondMY)
'            scrollTimer.Enabled = True
'        End If
    End If
End Sub

Private Sub p_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If scrollerObject Is infrm Then Call scrollerObject_MouseDown(Button, Shift, p(Index).Left + X, p(Index).top + Y)
End Sub
Private Sub p_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If scrollerObject Is infrm Then Call scrollerObject_MouseMove(Button, Shift, p(Index).Left + X, p(Index).top + Y)
End Sub
Private Sub p_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If scrollerObject Is infrm Then Call scrollerObject_MouseUp(Button, Shift, p(Index).Left + X, p(Index).top + Y)
End Sub

Private Sub calculateScroll(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        scrollTimer.Enabled = False
        Dim top As Long
        top = (scr.top - scrollRect.top) + (Y / 15) - firstMY
        scrValue = ((scrMax * top) / (scrollRect.Height - scr.Height))
        ScrollValue = scrValue
'        scr.top = scrollRect.top + scrValue
'        scrMid.top = scr.top + ((scr.Height / 2) - (scrMid.Height / 2))
'        scrTop.top = scr.top - scrTop.Height
'        scrDown.top = scr.top + scr.Height
        firstMY = Y / 15
    End If
End Sub
Private Sub scrRectEventer_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        Me.MousePointer = MousePointerConstants.vbSizeNS
        firstMY = Y / 15
        Call calculateScroll(Button, Shift, X, Y)
    ElseIf Button = 4 Then
        ScrollValue = 0
    End If
End Sub
Private Sub scrRectEventer_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call calculateScroll(Button, Shift, X, Y)
End Sub
Private Sub scrRectEventer_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        Me.MousePointer = MousePointerConstants.vbDefault
    End If
End Sub

Private Sub scr_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single): Call scrRectEventer_MouseDown(Button, Shift, CSng(scr.Left - scrRectEventer.Left) * 15 + X, CSng(scr.top - scrRectEventer.top) * 15 + Y): End Sub
Private Sub scr_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single): Call scrRectEventer_MouseMove(Button, Shift, CSng(scr.Left - scrRectEventer.Left) * 15 + X, CSng(scr.top - scrRectEventer.top) * 15 + Y): End Sub
Private Sub scr_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single): Call scrRectEventer_MouseUp(Button, Shift, CSng(scr.Left - scrRectEventer.Left) * 15 + X, CSng(scr.top - scrRectEventer.top) * 15 + Y): End Sub
Private Sub scrTop_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single): Call scrRectEventer_MouseDown(Button, Shift, CSng(scrTop.Left - scrRectEventer.Left) * 15 + X, CSng(scrTop.top - scrRectEventer.top) * 15 + Y): End Sub
Private Sub scrTop_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single): Call scrRectEventer_MouseMove(Button, Shift, CSng(scrTop.Left - scrRectEventer.Left) * 15 + X, CSng(scrTop.top - scrRectEventer.top) * 15 + Y): End Sub
Private Sub scrTop_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single): Call scrRectEventer_MouseUp(Button, Shift, CSng(scrTop.Left - scrRectEventer.Left) * 15 + X, CSng(scrTop.top - scrRectEventer.top) * 15 + Y): End Sub
Private Sub scrDown_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single): Call scrRectEventer_MouseDown(Button, Shift, CSng(scrDown.Left - scrRectEventer.Left) * 15 + X, CSng(scrDown.top - scrRectEventer.top) * 15 + Y): End Sub
Private Sub scrDown_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single): Call scrRectEventer_MouseMove(Button, Shift, CSng(scrDown.Left - scrRectEventer.Left) * 15 + X, CSng(scrDown.top - scrRectEventer.top) * 15 + Y): End Sub
Private Sub scrDown_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single): Call scrRectEventer_MouseUp(Button, Shift, CSng(scrDown.Left - scrRectEventer.Left) * 15 + X, CSng(scrDown.top - scrRectEventer.top) * 15 + Y): End Sub
Private Sub scrMid_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single): Call scrRectEventer_MouseDown(Button, Shift, CSng(scrMid.Left - scrRectEventer.Left) * 15 + X, CSng(scrMid.top - scrRectEventer.top) * 15 + Y): End Sub
Private Sub scrMid_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single): Call scrRectEventer_MouseMove(Button, Shift, CSng(scrMid.Left - scrRectEventer.Left) * 15 + X, CSng(scrMid.top - scrRectEventer.top) * 15 + Y): End Sub
Private Sub scrMid_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single): Call scrRectEventer_MouseUp(Button, Shift, CSng(scrMid.Left - scrRectEventer.Left) * 15 + X, CSng(scrMid.top - scrRectEventer.top) * 15 + Y): End Sub

'------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------
'--################################################################################--
'--################################################################################--
'------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------
'search

Private Sub searchText_GotFocus()
    If searchText.Text = DEFAULTSEARCHTEXT Then
        searchText.Text = ""
        searchText.ForeColor = &H404040
    End If
End Sub

Private Sub searchText_LostFocus()
    If searchText.Text = "" Then
        searchText.Text = DEFAULTSEARCHTEXT
        searchText.ForeColor = &HC0C0C0
    End If
End Sub

Private Sub searchText_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call SearchThis(searchText.Text)
    Else
        Call FilterThis(searchText.Text)
    End If
End Sub

'------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------
'--################################################################################--
'--################################################################################--
'------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------
'page manager

Private Sub SetBackEnabled(Value As Boolean)
    backButton(0).Enabled = Value
    If Value Then
        Set backButton(0).Picture = backButton(1).Picture
    Else
        Set backButton(0).Picture = backButton(2).Picture
    End If
End Sub
Private Sub SetNextEnabled(Value As Boolean)
    nextButton(0).Enabled = Value
    If Value Then
        Set nextButton(0).Picture = nextButton(1).Picture
    Else
        Set nextButton(0).Picture = nextButton(2).Picture
    End If
End Sub

Private Sub homeButton_Click()
    If scrollerObject Is infrm Then Exit Sub
    Call ScrollThis(infrm)
    Call infrm.ZOrder(ZOrderConstants.vbBringToFront)
End Sub
Private Sub mastersettingsButton_Click()
    If scrollerObject Is Settings.frm Then Exit Sub
    Call ScrollThis(Settings.frm)
    Call Settings.frm.ZOrder(ZOrderConstants.vbBringToFront)
End Sub

'------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------
'--################################################################################--
'--################################################################################--
'------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------
'search

Public Sub SearchThis(s As String)
    searchText.Text = s
End Sub
Public Sub FilterThis(s As String)
    searchText.Text = s
End Sub

'------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------
'--################################################################################--
'--################################################################################--
'------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------
'history

Private Sub backButton_Click(Index As Integer): Call HBack: End Sub
Private Sub nextButton_Click(Index As Integer): Call HNext: End Sub

Private Sub appendHistory(Obj As Object)
    Debug.Print "Executing Function appendHistory()"
    If historyCount < historyMax Then
        ReDim Preserve history(historyCount)
        history(historyCount) = Str
        historyCount = historyCount + 1
    Else
        Dim i As Long
        For i = 0 To historyCount - 2
            history(i) = history(i + 1)
        Next
        history(i) = Str
    End If
    Call SetBackEnabled(True)
    Call SetNextEnabled(False)
    Debug.Print "Execution Of Function appendHistory() Had Ended"
End Sub

'Commands:
'scrollthis
Private Sub HBack()
    Dim Obj As Object
    
End Sub
Public Sub HNext()
    
End Sub

Private Sub ShowInRect(Index As Long)
    If Not scrollerObject Is infrm Then Exit Sub
    If Not ScrollVisible Then Exit Sub
    Dim minY As Long, maxY As Long, cY As Long
    minY = (Index) * (p(0).Height / 15)
    maxY = minY - (frm.Height + (p(Index).Height / 15))
    cY = (ScrollValue * infrm.Height / 15) / ScrollMax
    'If cY < minY Then cY = minY: GoTo changed
    'If cY > maxY < 20 Then cY = maxY: GoTo changed
    Exit Sub
changed:
    ScrollValue = cY
End Sub

Private Function getSumOfHeights(ByVal fromIndex As Long, ByVal toIndex As Long) As Long
    If showingPlugins <= 0 Then Exit Function
    If fromIndex < 0 Then fromIndex = 0
    If toIndex < 0 Then toIndex = 0
    If fromIndex >= showingPlugins Then fromIndex = showingPlugins - 1
    If toIndex >= showingPlugins Then toIndex = showingPlugins - 1
    Dim sum As Long
    Dim i As Long
    For i = fromIndex To toIndex
        sum = sum + p(i).Height
    Next
    getSumOfHeights = sum
End Function

Private Sub p_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyUp Then
        If Index - 1 >= 0 Then Call p(Index - 1).SetFocus
    ElseIf KeyCode = vbKeyDown Then
        If Index + 1 < showingPlugins Then Call p(Index + 1).SetFocus
    End If
End Sub
Private Sub p_KeyPress(Index As Integer, KeyAscii As Integer)
    '
End Sub
Private Sub p_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    '
End Sub

Private Sub p_GotFocus(Index As Integer)
    Call ShowInRect(CLng(Index))
End Sub
'Private Sub p_LostFocus(Index As Integer)
    'Call ShowInRect(CLng(Index))
'End Sub
