VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About MintAPI"
   ClientHeight    =   7275
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5085
   BeginProperty Font 
      Name            =   "Consolas"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmUnLicensed.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7275
   ScaleWidth      =   5085
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox infoText 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   5295
      Left            =   1080
      MultiLine       =   -1  'True
      TabIndex        =   5
      Text            =   "frmUnLicensed.frx":2370A
      Top             =   1200
      Width           =   3735
   End
   Begin VB.Timer unregVis 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   0
      Top             =   0
   End
   Begin VB.CommandButton btnOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   390
      Left            =   3450
      TabIndex        =   3
      Top             =   6600
      Width           =   1290
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "by Ali Mousavi Kherad | alimousavikherad@gmail.com"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Left            =   0
      TabIndex        =   4
      Top             =   7050
      Visible         =   0   'False
      Width           =   5100
   End
   Begin VB.Label regState 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MintAPI Not Registered"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   1125
      TabIndex        =   2
      Top             =   825
      Visible         =   0   'False
      Width           =   1980
   End
   Begin VB.Label lblVersion 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0.0.0.2012 "
      ForeColor       =   &H00808080&
      Height          =   195
      Left            =   1125
      TabIndex        =   1
      Top             =   600
      Width           =   990
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
      ForeColor       =   &H00E89539&
      Height          =   360
      Left            =   1125
      TabIndex        =   0
      Top             =   300
      Width           =   1260
   End
   Begin VB.Image logo 
      Height          =   750
      Left            =   225
      Picture         =   "frmUnLicensed.frx":2390E
      Top             =   300
      Width           =   750
   End
   Begin VB.Image unregLogo 
      Height          =   750
      Left            =   225
      Picture         =   "frmUnLicensed.frx":2455C
      Top             =   300
      Visible         =   0   'False
      Width           =   750
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim isNotRegister As Boolean

Private Sub btnOK_Click(): Call Unload(Me): End Sub

Private Sub Form_Load()
    Me.Caption = Mtr("About MintAPI")
    regState.Caption = Mtr("MintAPI Not Registered")
    Dim str As String
    str = "MintAPI is a strong api provided" & vbCrLf & _
          "For atl - based apps such as VB6" & vbCrLf & _
          "Programmer:   Ali Mousavi Kherad" & vbCrLf & _
          "     alimousavikherad@ gmail.com" & vbCrLf & _
          vbCrLf & _
          "-Mail me if you have any idea or" & vbCrLf & _
          " Problem using MintAPI.         " & vbCrLf & _
          vbCrLf & _
          "-MintAPI provided under LGPL(v3)" & vbCrLf & _
          "  You can use it in anyway, but " & vbCrLf & _
          "  You must include my name.     " & vbCrLf & _
          vbCrLf & _
          "Thank you for using MintAPI.    " & vbCrLf & _
          vbCrLf & _
          ":) Ali Mousavi Kherad           "
'          "-Many of MintAPI modules is free" & vbCrLf & _
'          " To use but some little features" & vbCrLf & _
'          " Anyway if you want to use free " & vbCrLf & _
'          " Modules you must register your " & vbCrLf & _
'          " Copy at #######################" & vbCrLf & _
'          " To activate MintAPI , other way" & vbCrLf & _
'          " You may recieve errors on some " & vbCrLf & _
'          " Parts execution And unregister " & vbCrLf & _
'          " Form always will be open on lib" & vbCrLf & _
'          " Start."
    infoText.Text = Mtr(str)
    lblVersion.Caption = APP_VERSIONSTRING
    isNotRegister = False
End Sub

Private Function SetRegState() As Boolean
    On Error GoTo Err
    If Not RegisterationState Then
        regState.Visible = True
        logo.Visible = False
        unregLogo.Visible = True
        isNotRegister = True
        Me.Show
    Else
        isNotRegister = False
    End If
    SetRegState = True
    btnOK.Visible = isNotRegister
    Exit Function
Err:
    SetRegState = False
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If (Not RegisterationState) And (isNotRegister) Then Cancel = True
End Sub

Private Sub unregVis_Timer()
    If SetRegState Then
        unregVis.Enabled = False
    Else
        unregVis.Interval = 2000
    End If
End Sub
