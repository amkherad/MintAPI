VERSION 5.00
Begin VB.Form edt_Language_Settings 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Language File Settings..."
   ClientHeight    =   4605
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6240
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
   Icon            =   "edt_Language_Settings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4605
   ScaleWidth      =   6240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton btnOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   4830
      TabIndex        =   5
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CommandButton btnCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3510
      TabIndex        =   6
      Top             =   3960
      Width           =   1215
   End
   Begin VB.TextBox txtDescription 
      Height          =   1365
      Left            =   1665
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   2250
      Width           =   4395
   End
   Begin VB.TextBox txtRegion 
      Height          =   285
      Left            =   1665
      MaxLength       =   20
      TabIndex        =   2
      Top             =   1305
      Width           =   4395
   End
   Begin VB.TextBox txtShortName 
      Height          =   285
      Left            =   1665
      MaxLength       =   20
      TabIndex        =   1
      Top             =   810
      Width           =   4395
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   1665
      MaxLength       =   50
      TabIndex        =   0
      Top             =   315
      Width           =   4395
   End
   Begin VB.CheckBox chkRightToLeft 
      Caption         =   "Right To Left"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1665
      TabIndex        =   3
      Top             =   1800
      Width           =   4410
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Description:"
      ForeColor       =   &H00404040&
      Height          =   195
      Index           =   3
      Left            =   0
      TabIndex        =   10
      Top             =   2295
      Width           =   1620
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Region:"
      ForeColor       =   &H00404040&
      Height          =   195
      Index           =   2
      Left            =   0
      TabIndex        =   9
      Top             =   1350
      Width           =   1620
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Short Name (Key):"
      ForeColor       =   &H00404040&
      Height          =   195
      Index           =   1
      Left            =   0
      TabIndex        =   8
      Top             =   855
      Width           =   1620
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Name:"
      ForeColor       =   &H00404040&
      Height          =   195
      Index           =   0
      Left            =   0
      TabIndex        =   7
      Top             =   360
      Width           =   1620
   End
End
Attribute VB_Name = "edt_Language_Settings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim p As Object

Dim p_modified As Boolean

Friend Sub SetParent(Parent As Object)
    Set p = Parent
End Sub

Private Sub btnCancel_Click()
    Call Unload(Me)
End Sub

Private Sub btnOK_Click()
    If Not p Is Nothing Then _
        Call p.SetEditorProperties(p_modified, txtName.Text, txtShortName.Text, txtRegion.Text, txtDescription.Text, chkRightToLeft.Value)

    Call Unload(Me)
End Sub

Private Sub chkRightToLeft_Click()
    p_modified = True
End Sub

Private Sub txtDescription_Change()
    p_modified = True
End Sub

Private Sub txtName_Change()
    p_modified = True
End Sub

Private Sub txtRegion_Change()
    p_modified = True
End Sub

Private Sub txtShortName_Change()
    p_modified = True
End Sub
