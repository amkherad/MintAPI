VERSION 5.00
Begin VB.Form toolsForm 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "MintAPI Tools"
   ClientHeight    =   9195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5265
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
   Icon            =   "tools_mForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9195
   ScaleWidth      =   5265
   StartUpPosition =   2  'CenterScreen
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.Timer tmEffect 
      Interval        =   700
      Left            =   4005
      Top             =   135
   End
   Begin VB.CommandButton btnSourceCodeManager 
      Caption         =   "Open &Source Code Manager"
      Height          =   615
      Left            =   495
      TabIndex        =   10
      Top             =   6615
      Width           =   2775
   End
   Begin VB.CommandButton btnHelp 
      Cancel          =   -1  'True
      Caption         =   "&Help"
      Height          =   375
      Left            =   225
      TabIndex        =   7
      Top             =   8595
      Width           =   855
   End
   Begin VB.CommandButton btnClose 
      Caption         =   "Clo&se"
      Height          =   375
      Left            =   3915
      TabIndex        =   4
      Top             =   8595
      Width           =   1105
   End
   Begin VB.CommandButton btnConfigurationEditor 
      Caption         =   "Open &Configuration Editor"
      Height          =   615
      Left            =   480
      TabIndex        =   1
      Top             =   4200
      Width           =   2775
   End
   Begin VB.CommandButton btnLanguageEditor 
      Caption         =   "Open &Language Editor"
      Height          =   615
      Left            =   480
      TabIndex        =   0
      Top             =   2400
      Width           =   2775
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   $"tools_mForm.frx":1082
      ForeColor       =   &H00404040&
      Height          =   1365
      Left            =   495
      TabIndex        =   9
      Top             =   5085
      Width           =   4335
   End
   Begin VB.Label lblMessage 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H0000F0C6&
      Height          =   255
      Left            =   360
      TabIndex        =   8
      Top             =   8100
      Width           =   4560
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   $"tools_mForm.frx":11BC
      ForeColor       =   &H00404040&
      Height          =   735
      Left            =   480
      TabIndex        =   6
      Top             =   3240
      Width           =   4335
   End
   Begin VB.Label langInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "1. Click on the button bellow to open the Language Editor to create language files for Your application."
      ForeColor       =   &H00404040&
      Height          =   735
      Left            =   480
      TabIndex        =   5
      Top             =   1680
      Width           =   4335
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Height          =   6975
      Left            =   240
      Top             =   1440
      Width           =   4780
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tools 0.1"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   225
      Left            =   1200
      TabIndex        =   3
      Top             =   840
      Width           =   945
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MintAPI"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   360
      Left            =   1200
      TabIndex        =   2
      Top             =   480
      Width           =   1260
   End
   Begin VB.Image Image1 
      Height          =   750
      Left            =   360
      Picture         =   "tools_mForm.frx":1245
      Top             =   360
      Width           =   750
   End
End
Attribute VB_Name = "toolsForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const MESSAGE As String = "MintAPI Tools version 1.0 [part of MintAPI2ndLayer] by Ali Mousavi Kherad (alimousavikherad@gmail.com)"
Private Const MESSAGEC As Long = 50 'characters.

Dim lstOpened As New Collection
Attribute lstOpened.VB_VarHelpID = -1

Dim MessageShowen_Start As Long

Private Sub btnClose_Click()
    Call Unload(Me)
End Sub

Private Sub btnConfigurationEditor_Click()
On Error GoTo err
    If Not LayerAPI.IsInModalState Then
        Dim cfE As Object
        Set cfE = ConfigurationEditor
        Call lstOpened.Add(cfE)
        Call cfE.Show
        Call cfE.Focus
    Else
        Call MsgBox("Another form showing modally.")
    End If
    Exit Sub
err:
    ShowErrorMessage (err.Description)
End Sub

Private Sub btnHelp_Click()
    '
End Sub

Private Sub btnLanguageEditor_Click()
On Error GoTo err
    If Not LayerAPI.IsInModalState Then
        Dim lnE As Object
        Set lnE = LanguageEditor
        Call lstOpened.Add(lnE)
        Call lnE.Show
        Call lnE.Focus
    Else
        Call MsgBox("Another form showing modally.")
    End If
    Exit Sub
err:
    ShowErrorMessage (err.Description)
End Sub

Private Sub btnSourceCodeManager_Click()
On Error GoTo err
    If Not LayerAPI.IsInModalState Then
        Dim scM As Object
        Set scM = SourceCodeManager
        Call lstOpened.Add(scM)
        Call scM.Show
        Call scM.Focus
    Else
        Call MsgBox("Another form showing modally.")
    End If
    Exit Sub
err:
    ShowErrorMessage (err.Description)
End Sub

Private Sub Form_Load()
    lblMessage.Caption = Left(MESSAGE, MESSAGEC)
    MessageShowen_Start = 0
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
    Dim i As Long
    For i = 1 To lstOpened.Count
        If Not lstOpened(i) Is Nothing Then _
            Call lstOpened(i).CloseForm
    Next
    Exit Sub
err:
    ShowErrorMessage (err.Description)
End Sub

Private Sub tmEffect_Timer()
    lblMessage.Caption = Mid(MESSAGE & " " & MESSAGE, (MessageShowen_Start Mod (Len(MESSAGE) + 1)) + 1, MESSAGEC)
    MessageShowen_Start = MessageShowen_Start + 1
End Sub
