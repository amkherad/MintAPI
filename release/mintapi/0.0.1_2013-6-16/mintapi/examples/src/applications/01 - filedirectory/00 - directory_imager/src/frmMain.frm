VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H009A734B&
   BorderStyle     =   0  'None
   Caption         =   "Directory Imager"
   ClientHeight    =   7410
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8355
   BeginProperty Font 
      Name            =   "Consolas"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   494
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   557
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H009A734B&
      Caption         =   "Compare"
      ForeColor       =   &H0000FFFF&
      Height          =   2535
      Left            =   315
      TabIndex        =   10
      Top             =   4140
      Width           =   7710
      Begin VB.CommandButton compareBtn 
         Caption         =   "Compare"
         Height          =   375
         Left            =   6480
         TabIndex        =   15
         Top             =   1935
         Width           =   1005
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H009A734B&
         Caption         =   "Check file times (create date/time - last modify)"
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   225
         TabIndex        =   14
         Top             =   1305
         Width           =   4830
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H009A734B&
         Caption         =   "Accept newly added files"
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   225
         TabIndex        =   13
         Top             =   990
         Width           =   4560
      End
      Begin VB.TextBox dirC 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Height          =   285
         Left            =   225
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   630
         Width           =   7035
      End
      Begin VB.Image dirCBrowse 
         Height          =   240
         Left            =   7290
         Picture         =   "frmMain.frx":1082
         Top             =   630
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Select Directory Image File..."
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   225
         TabIndex        =   11
         Top             =   405
         Width           =   2700
      End
   End
   Begin VB.Frame frmCreate 
      BackColor       =   &H009A734B&
      Caption         =   "Create"
      ForeColor       =   &H0000FFFF&
      Height          =   2355
      Left            =   315
      TabIndex        =   2
      Top             =   1710
      Width           =   7710
      Begin VB.TextBox dirP 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Height          =   285
         Left            =   225
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   630
         Width           =   7035
      End
      Begin VB.CommandButton createBtn 
         Caption         =   "Create"
         Height          =   375
         Left            =   6480
         TabIndex        =   9
         Top             =   1800
         Width           =   1005
      End
      Begin VB.TextBox inc_formats 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Height          =   465
         Left            =   225
         MultiLine       =   -1  'True
         TabIndex        =   7
         Text            =   "frmMain.frx":12FF
         Top             =   1215
         Width           =   3570
      End
      Begin VB.TextBox ex_formats 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Height          =   465
         Left            =   3915
         MultiLine       =   -1  'True
         TabIndex        =   8
         Text            =   "frmMain.frx":1307
         Top             =   1215
         Width           =   3570
      End
      Begin VB.Image dirPBrowse 
         Height          =   240
         Left            =   7290
         Picture         =   "frmMain.frx":1318
         Top             =   630
         Width           =   240
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Select a Path to Save Image..."
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   225
         TabIndex        =   3
         Top             =   405
         Width           =   2700
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Excludes:"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   3915
         TabIndex        =   6
         Top             =   990
         Width           =   810
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Includes:"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   225
         TabIndex        =   5
         Top             =   990
         Width           =   810
      End
   End
   Begin VB.TextBox dir 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      Height          =   285
      Left            =   315
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   1305
      Width           =   7440
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Current Action:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   225
      TabIndex        =   20
      Top             =   6795
      Width           =   1350
   End
   Begin VB.Label log 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "..."
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1575
      TabIndex        =   19
      Top             =   6795
      Width           =   270
   End
   Begin VB.Image minimizeBTNNH 
      Height          =   240
      Left            =   7740
      Picture         =   "frmMain.frx":1595
      Top             =   540
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image closeBTNNH 
      Height          =   240
      Left            =   7965
      Picture         =   "frmMain.frx":15EB
      Top             =   540
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image minimizeBTNH 
      Height          =   240
      Left            =   7740
      Picture         =   "frmMain.frx":1688
      Top             =   360
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image closeBTNH 
      Height          =   240
      Left            =   7965
      Picture         =   "frmMain.frx":3A5A
      Top             =   360
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   3
      X1              =   21
      X2              =   531
      Y1              =   54
      Y2              =   54
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Select Directory"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   315
      TabIndex        =   0
      Top             =   1080
      Width           =   1440
   End
   Begin VB.Image dirBrowse 
      Height          =   240
      Left            =   7785
      Picture         =   "frmMain.frx":5E2C
      Top             =   1320
      Width           =   240
   End
   Begin VB.Image icon 
      Height          =   360
      Left            =   135
      Picture         =   "frmMain.frx":60A9
      Top             =   225
      Width           =   360
   End
   Begin VB.Shape pg 
      BorderColor     =   &H000080FF&
      BorderWidth     =   5
      Height          =   6495
      Left            =   45
      Top             =   30
      Width           =   8295
   End
   Begin VB.Image closeBTN 
      Height          =   240
      Left            =   7965
      Picture         =   "frmMain.frx":65C8
      Top             =   135
      Width           =   240
   End
   Begin VB.Image minimizeBTN 
      Height          =   240
      Left            =   7740
      Picture         =   "frmMain.frx":6665
      Top             =   135
      Width           =   240
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "by Ali Mousavi Kherad (alimousavikherad@gmail.com)  |  09194895618"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   225
      TabIndex        =   16
      Top             =   7020
      Width           =   5940
   End
   Begin VB.Label lblMove 
      BackStyle       =   0  'Transparent
      Height          =   825
      Left            =   -90
      TabIndex        =   18
      Top             =   0
      Width           =   8925
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Directory Imager"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   540
      TabIndex        =   17
      Top             =   315
      Width           =   1440
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function API_CreateRoundRectRgn Lib "gdi32" Alias "CreateRoundRectRgn" (ByVal x1 As Integer, ByVal y1 As Integer, ByVal x2 As Integer, ByVal y2 As Integer, ByVal x3 As Integer, ByVal y3 As Integer) As Long
Private Declare Function API_DeleteObject Lib "gdi32" Alias "DeleteObject" (ByVal hObject As Long) As Long
Private Declare Function API_SetWindowRgn Lib "user32" Alias "SetWindowRgn" (ByVal hwnd As Long, ByVal hRGN As Long, ByVal bRedraw As Boolean) As Long

Const FORMWIDTH As Long = 555
Const FORMHEIGHT As Long = 494

Dim mdX As Long, mdY As Long

Dim cwp As Long

Private Sub closeBTN_Click(): Call Unload(Me): End Sub

Private Sub Form_Load()
    Call Me.Move(Me.Left, Me.Top, FORMWIDTH * 15, FORMHEIGHT * 15)
    Dim hRGN As Long
    hRGN = API_CreateRoundRectRgn(0, 0, FORMWIDTH, FORMHEIGHT, 5, 5)
    Call API_SetWindowRgn(hwnd, hRGN, True)
    Call API_DeleteObject(hRGN)
    log.Caption = "..."
End Sub

Private Sub icon_DblClick(): Call Unload(Me): End Sub

Private Sub lblMove_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        mdX = X
        mdY = Y
    End If
End Sub

Private Sub lblMove_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        Call Me.Move(Me.Left + X - mdX, Me.Top + Y - mdY)
    End If
    If cwp = 1 Then _
        Set minimizeBTN.Picture = minimizeBTNNH.Picture
    If cwp = 2 Then _
        Set closeBTN.Picture = closeBTNNH.Picture
    cwp = 0
End Sub

Private Sub minimizeBTN_Click(): Me.WindowState = 1: End Sub
Private Sub Form_Resize()
    On Error Resume Next
    Call pg.Move(0, 0, ScaleWidth - 1, ScaleHeight - 2)
End Sub

Private Sub minimizeBTN_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set minimizeBTN.Picture = minimizeBTNH.Picture
    Set closeBTN.Picture = closeBTNNH.Picture
    cwp = 1
End Sub
Private Sub closeBTN_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set closeBTN.Picture = closeBTNH.Picture
    Set minimizeBTN.Picture = minimizeBTNNH.Picture
    cwp = 2
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If cwp = 1 Then _
        Set minimizeBTN.Picture = minimizeBTNNH.Picture
    If cwp = 2 Then _
        Set closeBTN.Picture = closeBTNNH.Picture
    cwp = 0
End Sub



Private Sub createBtn_Click()
    Call StartImaging(dir.Text, dirP.Text)
End Sub




Private Sub dirBrowse_Click()
On Error GoTo err
    dir.Text = Directory.ChooseDirectory(Me.hwnd, "Select a directory to image from...").AbsolutePath
err:
End Sub
Private Sub dirCBrowse_Click()
On Error GoTo err
    dirC.Text = File.ChooseFile(Me.hwnd, fdOpen, "Select a directory image file...", "All Files(*.*)", "", CurrentDirectory).AbsolutePath
err:
End Sub
Private Sub dirPBrowse_Click()
On Error GoTo err
    dirP.Text = File.ChooseFile(Me.hwnd, fdSave, "Select a path to save image...", "All Files(*.*)", "", CurrentDirectory).AbsolutePath
err:
End Sub
