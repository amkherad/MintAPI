VERSION 5.00
Begin VB.Form mForm 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "miccy Picture Tools Advanced Version"
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
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   Picture         =   "mForm.frx":0000
   ScaleHeight     =   566
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   840
   StartUpPosition =   2  'CenterScreen
   WhatsThisHelp   =   -1  'True
   Begin VB.Timer scrollTimer 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   7575
      Top             =   150
   End
   Begin VB.ComboBox sorter 
      Height          =   315
      ItemData        =   "mForm.frx":8DB2
      Left            =   6075
      List            =   "mForm.frx":8DDA
      Style           =   2  'Dropdown List
      TabIndex        =   16
      Top             =   885
      Width           =   2490
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
      Left            =   8775
      TabIndex        =   14
      Text            =   "Search Plugins Here"
      Top             =   900
      Width           =   3045
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
      Begin VB.Frame plgfrm 
         Appearance      =   0  'Flat
         BackColor       =   &H00FBD2AE&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3390
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
   Begin VB.Image scrMid 
      Height          =   105
      Left            =   12180
      Picture         =   "mForm.frx":8EAC
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
      Picture         =   "mForm.frx":905C
      Stretch         =   -1  'True
      Top             =   960
      Width           =   135
   End
   Begin VB.Image scrDown 
      Height          =   60
      Left            =   12180
      Picture         =   "mForm.frx":91AB
      Top             =   8145
      Width           =   135
   End
   Begin VB.Image scrTop 
      Height          =   60
      Left            =   12180
      Picture         =   "mForm.frx":931A
      Top             =   900
      Width           =   135
   End
   Begin VB.Image openPlugins 
      Height          =   240
      Index           =   3
      Left            =   5550
      Picture         =   "mForm.frx":9489
      ToolTipText     =   "List All Open Plugins"
      Top             =   900
      Width           =   240
   End
   Begin VB.Image backButton 
      Height          =   240
      Index           =   2
      Left            =   4800
      Picture         =   "mForm.frx":9568
      ToolTipText     =   "<Previeus Page"
      Top             =   450
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image nextButton 
      Height          =   240
      Index           =   2
      Left            =   5025
      Picture         =   "mForm.frx":95F7
      ToolTipText     =   "Next Page>"
      Top             =   450
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image backButton 
      Height          =   240
      Index           =   1
      Left            =   4800
      Picture         =   "mForm.frx":9686
      ToolTipText     =   "<Previeus Page"
      Top             =   600
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image nextButton 
      Height          =   240
      Index           =   1
      Left            =   5025
      Picture         =   "mForm.frx":9715
      ToolTipText     =   "Next Page>"
      Top             =   600
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image homeButton 
      Height          =   240
      Left            =   5325
      Picture         =   "mForm.frx":97A4
      ToolTipText     =   "Plugin Picker Page"
      Top             =   900
      Width           =   240
   End
   Begin VB.Image rssButton 
      Height          =   240
      Left            =   5775
      Picture         =   "mForm.frx":984F
      Top             =   900
      Width           =   240
   End
   Begin VB.Image nextButton 
      Height          =   240
      Index           =   0
      Left            =   5025
      Picture         =   "mForm.frx":9AAA
      ToolTipText     =   "Next Page>"
      Top             =   900
      Width           =   240
   End
   Begin VB.Image backButton 
      Height          =   240
      Index           =   0
      Left            =   4800
      Picture         =   "mForm.frx":9B39
      ToolTipText     =   "<Previeus Page"
      Top             =   900
      Width           =   240
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   285
      Left            =   8700
      Top             =   900
      Width           =   240
   End
   Begin VB.Image searchGo 
      Height          =   270
      Left            =   11850
      Picture         =   "mForm.frx":9BC8
      Stretch         =   -1  'True
      Top             =   900
      Width           =   240
   End
   Begin VB.Image search 
      Height          =   420
      Left            =   8700
      Picture         =   "mForm.frx":9CA7
      Stretch         =   -1  'True
      Top             =   825
      Width           =   3315
   End
   Begin VB.Image searchRight 
      Height          =   420
      Left            =   12000
      Picture         =   "mForm.frx":9E13
      Top             =   825
      Width           =   90
   End
   Begin VB.Image searchLeft 
      Height          =   420
      Left            =   8625
      Picture         =   "mForm.frx":A02B
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
      Left            =   1050
      Picture         =   "mForm.frx":A23C
      Top             =   825
      Width           =   180
   End
   Begin VB.Image menuhLeft 
      Height          =   420
      Left            =   375
      Picture         =   "mForm.frx":A4F9
      Top             =   825
      Width           =   180
   End
   Begin VB.Image menuh 
      Height          =   420
      Left            =   525
      Picture         =   "mForm.frx":A79F
      Stretch         =   -1  'True
      Top             =   825
      Width           =   540
   End
   Begin VB.Image menu 
      Height          =   420
      Left            =   300
      Picture         =   "mForm.frx":A937
      Stretch         =   -1  'True
      Top             =   825
      Width           =   4365
   End
   Begin VB.Image menuRight 
      Height          =   420
      Left            =   4650
      Picture         =   "mForm.frx":AAA3
      Top             =   825
      Width           =   90
   End
   Begin VB.Image menuLeft 
      Height          =   420
      Left            =   225
      Picture         =   "mForm.frx":ACBB
      Top             =   825
      Width           =   90
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
   Begin VB.Image closeBTN 
      Height          =   240
      Left            =   12150
      Picture         =   "mForm.frx":AECC
      Top             =   150
      Width           =   240
   End
   Begin VB.Image minimizeBTN 
      Height          =   240
      Left            =   11925
      Picture         =   "mForm.frx":D29E
      Top             =   150
      Width           =   240
   End
   Begin VB.Label title 
      BackStyle       =   0  'Transparent
      Height          =   765
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   12615
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
   Begin VB.Label titleCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "miccy Ultimate Tools"
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
      Width           =   3000
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
   Begin VB.Label scrRectEventer 
      BackStyle       =   0  'Transparent
      Height          =   7440
      Left            =   12120
      TabIndex        =   17
      Top             =   840
      Width           =   270
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

Const SCROLLMININC As Long = 5
Dim scrMax As Long, scrValue As Long
Dim scrEfcTop As Double, scrConTop As Double, scrConIncrease As Double
Dim scrEfcTopIncrease As Double
Dim scrollerObject As Object
Attribute scrollerObject.VB_VarHelpID = -1
Dim scrMDY As Long

'------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------
'--################################################################################--
'--################################################################################--
'------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------
'main form

Private Sub Form_Load()
    mnuEfcLeftIncrease = 0: mnuEfcWidthIncrease = 0
    Call Me.Move(Me.Left, Me.Top, FORMWIDTH * 15, FORMHEIGHT * 15)
    Dim hRGN As Long
    hRGN = API_CreateRoundRectRgn(0, 0, FORMWIDTH, FORMHEIGHT, 16, 16)
    Call API_SetWindowRgn(hwnd, hRGN, True)
    Call API_DeleteObject(hRGN)
    
    mnuTimer.Interval = MNUEFFECTTIMERINTERVAL
    scrollTimer.Interval = SCREFFECTTIMERINTERVAL
    
    sorter.ListIndex = 0
    
    frm.Width = scrollRect.Left - frm.Left
    
    Call InitializeMenu
    Call InitializePlugins
    searchText.Text = DEFAULTSEARCHTEXT
    
    Call ScrollThis(infrm)
End Sub

Private Sub title_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single): fX = X: fY = Y: End Sub
Private Sub title_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        Call Me.Move(Me.Left + X - fX, Me.Top + Y - fY)
    End If
    Call mnuGoto(0)
End Sub
Private Sub closeBTN_Click(): Call Unload(Me): End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer): Call EndApp: End Sub

'Private Sub infrm_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single): Call mnuGoto(0): End Sub
'Private Sub frm_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single): Call mnuGoto(0): End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single): Call mnuGoto(0): End Sub

Private Sub minimizeBTN_Click(): Me.WindowState = 1: End Sub

Private Sub InitializePlugins()
    Call InitializePluginGUIItems
End Sub
Private Sub InitializePluginGUIItems()
    plgfrm.Top = frm.Top
    infrm.Top = frm.Top
    plgfrm.Width = frm.Width * 15
    plgfrm.Height = frm.Height * 15
    infrm.Width = plgfrm.Width
    
    p(0).Width = infrm.Width
    noplugin.Width = infrm.Width
    noplugin.Top = ((frm.Height * 15) / 2) - (noplugin.Height / 2)
    
    Dim i As Long, maxHeight As Long
    maxHeight = p(0).Height
    For i = 1 To 100
        Call Load(p(i))
        p(i).Top = p(i - 1).Top + p(i - 1).Height
        p(i).Visible = True
        maxHeight = maxHeight + p(i).Height
        p(i).Caption = "Plugin Name " & i
    Next
    infrm.Height = maxHeight
End Sub
Private Sub InitializeMenu()
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
    mTop = Me.ScaleTop + mnu(Index).Top + mnu(Index).Height + EXTRAMENUTOP
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

Private Sub p_LeftClick(Index As Integer, X As Long, Y As Long)
    'scrMDY = Y
End Sub
Private Sub p_StartScrolling(Index As Integer, ToTop As Boolean, Power As Double)
    scrConIncrease = IIf(ToTop, -Power, Power)
    scrollTimer.Enabled = True
End Sub
Private Sub p_magneticScrolling(Index As Integer, Value As Long)
    scrollTimer.Enabled = False
'    If infrm.top + Value < frm.Height And infrm.top + Value + infrm.Height > 0 Then
'        infrm.top = infrm.top + Value
'    End If
    ScrollValue = ScrollValue - Value
End Sub
Private Sub p_RightClick(Index As Integer, X As Long, Y As Long)
    '
End Sub

Private Property Get ScrollVisible() As Boolean
    ScrollVisible = scr.Visible
End Property
Private Property Let ScrollVisible(Value As Boolean)
    scr.Visible = Value
    scrTop.Visible = Value
    scrDown.Visible = Value
    scrMid.Visible = Value
    ScrollValue = scrValue
End Property

Private Sub ScrollThis(ByVal Obj As Object)
    Set scrollerObject = Obj
    Obj.Top = 0
    Dim maxim As Long
    maxim = (Obj.Height / 15) - frm.Height
    If maxim <= 0 Then
        ScrollVisible = False
    Else
        If Not ScrollVisible Then ScrollVisible = True
        ScrollMax = maxim
        ScrollValue = 0
    End If
End Sub
Private Property Get ScrollMax() As Long
    ScrollMax = scrMax
End Property
Private Property Let ScrollMax(Value As Long)
    If Value <= 0 Then Value = 1
    If Value >= 1000000 Then Value = 1000000
    scrMax = Value
    Dim bufHeight As Long
    bufHeight = scrollRect.Height - (Value / SCROLLMININC)
    If bufHeight < scrMid.Height + 2 Then bufHeight = scrMid.Height + 2
    scr.Height = bufHeight
    If scrValue > scrMax Then scrValue = scrMax
    If Not ScrollVisible Then Exit Property
    'ScrollValue = scrValue
End Property
Private Property Get ScrollValue() As Long
    ScrollValue = scrValue
End Property
Private Property Let ScrollValue(Value As Long)
    If Value < 0 Then Value = 0
    If Value > scrMax Then Value = scrMax
    scrValue = Value
    If Not ScrollVisible Then Exit Property
'    Call ScrollTo(Value)
    If Not scrollerObject Is Nothing Then scrollerObject.Top = -(scrValue * 15)
    scr.Top = scrollRect.Top + (((scrollRect.Height - scr.Height) * scrValue) / scrMax)
    scrMid.Top = scr.Top + ((scr.Height / 2) - (scrMid.Height / 2))
    scrTop.Top = scr.Top - scrTop.Height
    scrDown.Top = scr.Top + scr.Height
End Property
'Private Sub ScrollTo(ByVal Value As Long, Optional smooth As Boolean = True)
''    Dim bufTop As Long
''    bufTop = scrollRect.top + (((scrollRect.Height - scr.Height) * scrValue) / scrMax)
'''    If smooth Then
''        scrEfcTop = scr.top
''        Const DINGS As Double = SCREFFECTLATENCY / SCREFFECTTIMERINTERVAL
''        scrEfcTopIncrease = (bufTop - scrEfcTop) / DINGS
''        If Not (CLng(scrEfcTop) = bufTop) Then
''            scrollTimer.Enabled = True
''        End If
'''    Else
'''        scr.top = bufTop
'''        Call Scrolled(bufTop)
'''        scrMid.top = scr.top + ((scr.Height / 2) - (scrMid.Height / 2))
'''        scrTop.top = scr.top - scrTop.Height
'''        scrDown.top = scr.top + scr.Height
'''    End If
'End Sub
Private Sub scrollTimer_Timer()
    scrConTop = scrConTop + scrConIncrease
    ScrollValue = scrConTop
    If ScrollValue = 0 Or ScrollValue = scrollerObject.Height Then scrollTimer.Enabled = False
'    Dim lastTop As Long, bufTop As Long
'    lastTop = scrEfcTop
'    scrEfcTop = scrEfcTop + scrEfcTopIncrease
'    scrEfcTopIncrease = scrEfcTopIncrease / 1.105
'    'If scrEfcTopIncrease < 0.7 Then scrEfcTopIncrease = 0.7
'    bufTop = scrollRect.top + (((scrollRect.Height - scr.Height) * scrValue) / scrMax)
'    If lastTop >= bufTop Then
'        If CLng(scrEfcTop) <= bufTop Then
'            GoTo StopIt
'        End If
'    End If
'    If lastTop <= bufTop Then
'        If CLng(scrEfcTop) >= bufTop Then
'            GoTo StopIt
'        End If
'    End If
'    GoTo NoStopIt
'StopIt:
'    scrEfcTop = bufTop
'    scrollTimer.Enabled = False
'NoStopIt:
'    scr.top = CLng(scrEfcTop) ' + mnuPADDINGWIDTH_2
'    Call Scrolled(scrValue)
'    scrMid.top = scr.top + ((scr.Height / 2) - (scrMid.Height / 2))
'    scrTop.top = scr.top - scrTop.Height
'    scrDown.top = scr.top + scr.Height
End Sub
'Private Sub Scrolled(tempValue As Long) 'slot()
''    tempValue = tempValue - scrollRect.top
''    'tempValue = getScrollValueFromSCRTop(tempValue)
''    If Not scrollerObject Is Nothing Then scrollerObject.top = -(tempValue * 15)
'End Sub
'Private Function getScrollValueFromSCRTop(scrTop As Long) As Long
'    getScrollValueFromSCRTop = (scrMax * scrTop) / (scrollRect.Height - scr.Height)
'End Function
'Private Sub calculateScroll(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    If Button = 1 Then
'        Dim top As Long, bufRect As Long
'        top = (Y / 15)
'        bufRect = scrollRect.Height - (scr.Height / 2)
'        If top < 0 Then top = 0
'        If top > scrRectEventer.Height Then top = scrRectEventer.Height
'        scrValue = getScrollValueFromSCRTop((scr.top + top) - scrMDY)
'        titleCaption.Caption = scrValue
'        Call ScrollTo(scrValue, False)
'    End If
'End Sub
'Private Sub scrRectEventer_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    scrMDY = Y / 15
'    Call calculateScroll(Button, Shift, X, Y)
'End Sub
'Private Sub scrRectEventer_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    Call calculateScroll(Button, Shift, X, Y)
'    scrMDY = Y / 15
'End Sub
'Private Sub scr_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single): Call scrRectEventer_MouseDown(Button, Shift, CSng(scr.Left - scrRectEventer.Left) * 15 + X, CSng(scr.top - scrRectEventer.top) * 15 + Y): End Sub
'Private Sub scr_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single): Call scrRectEventer_MouseMove(Button, Shift, CSng(scr.Left - scrRectEventer.Left) * 15 + X, CSng(scr.top - scrRectEventer.top) * 15 + Y): End Sub
'Private Sub scrTop_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single): Call scrRectEventer_MouseDown(Button, Shift, CSng(scrTop.Left - scrRectEventer.Left) * 15 + X, CSng(scrTop.top - scrRectEventer.top) * 15 + Y): End Sub
'Private Sub scrTop_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single): Call scrRectEventer_MouseMove(Button, Shift, CSng(scrTop.Left - scrRectEventer.Left) * 15 + X, CSng(scrTop.top - scrRectEventer.top) * 15 + Y): End Sub
'Private Sub scrDown_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single): Call scrRectEventer_MouseDown(Button, Shift, CSng(scrDown.Left - scrRectEventer.Left) * 15 + X, CSng(scrDown.top - scrRectEventer.top) * 15 + Y): End Sub
'Private Sub scrDown_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single): Call scrRectEventer_MouseMove(Button, Shift, CSng(scrDown.Left - scrRectEventer.Left) * 15 + X, CSng(scrDown.top - scrRectEventer.top) * 15 + Y): End Sub
'Private Sub scrMid_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single): Call scrRectEventer_MouseDown(Button, Shift, CSng(scrMid.Left - scrRectEventer.Left) * 15 + X, CSng(scrMid.top - scrRectEventer.top) * 15 + Y): End Sub
'Private Sub scrMid_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single): Call scrRectEventer_MouseMove(Button, Shift, CSng(scrMid.Left - scrRectEventer.Left) * 15 + X, CSng(scrMid.top - scrRectEventer.top) * 15 + Y): End Sub

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
    End If
End Sub

Private Sub searchText_LostFocus()
    If searchText.Text = "" Then
        searchText.Text = DEFAULTSEARCHTEXT
    End If
End Sub
