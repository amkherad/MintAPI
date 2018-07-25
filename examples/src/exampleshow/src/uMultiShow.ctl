VERSION 5.00
Begin VB.UserControl uMultiShow 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   8685
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9615
   ScaleHeight     =   579
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   641
   ToolboxBitmap   =   "uMultiShow.ctx":0000
   Begin VB.Frame frm 
      BackColor       =   &H00F8D5BA&
      BorderStyle     =   0  'None
      Height          =   5955
      Left            =   90
      TabIndex        =   3
      Top             =   90
      Width           =   5640
      Begin VB.VScrollBar sc 
         Height          =   5865
         Left            =   5400
         TabIndex        =   5
         Top             =   45
         Width           =   195
      End
      Begin VB.Frame infrm 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   5820
         Left            =   45
         TabIndex        =   4
         Top             =   45
         Width           =   5280
         Begin examples.uMultiShow_ITM uMultiShow_ITMs 
            Height          =   1050
            Left            =   45
            TabIndex        =   6
            Top             =   90
            Width           =   5190
            _ExtentX        =   9155
            _ExtentY        =   1852
         End
      End
   End
   Begin VB.Frame frmD 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   5955
      Left            =   5895
      TabIndex        =   0
      Top             =   90
      Width           =   2445
      Begin VB.Label lblCaption 
         BackStyle       =   0  'Transparent
         Caption         =   "TITLE..."
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000F0C6&
         Height          =   870
         Left            =   45
         TabIndex        =   2
         Top             =   45
         Width           =   2370
      End
      Begin VB.Label lblDescription 
         BackStyle       =   0  'Transparent
         Caption         =   "Content..."
         ForeColor       =   &H00F8D5BA&
         Height          =   4965
         Left            =   45
         TabIndex        =   1
         Top             =   945
         Width           =   2370
      End
   End
   Begin VB.Shape frmBorder 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   5  'Downward Diagonal
      Height          =   6045
      Left            =   45
      Top             =   45
      Width           =   5730
   End
   Begin VB.Shape frmDBorder 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   5  'Downward Diagonal
      Height          =   6045
      Left            =   5850
      Top             =   45
      Width           =   3300
   End
End
Attribute VB_Name = "uMultiShow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Sub sc_Validate(Cancel As Boolean)
    Cancel = True
End Sub

Private Sub UserControl_Resize()
Const SBORDER As Long = 2
    On Error Resume Next
    Call frmBorder.Move(frmBorder.Left, frmBorder.Top, ScaleWidth - frmBorder.Left * 3 - frmDBorder.Width, ScaleHeight - frmBorder.Top * 2)
    Call frm.Move(frmBorder.Left + SBORDER, frmBorder.Top + SBORDER, frmBorder.Width - SBORDER * 2, frmBorder.Height - SBORDER * 2)
    Call frmDBorder.Move(frmBorder.Width + frmBorder.Left * 2, frmDBorder.Top, frmDBorder.Width, ScaleHeight - frmDBorder.Top * 2)
    Call frmD.Move(frmDBorder.Left + SBORDER, frmDBorder.Top + SBORDER, frmDBorder.Width - SBORDER * 2, frmDBorder.Height - SBORDER * 2)
    Call lblDescription.Move(lblDescription.Left, lblDescription.Top, frmD.Width - lblDescription.Left * 2, frmD.Height - lblDescription.Top - lblDescription.Left)
End Sub
