VERSION 5.00
Begin VB.Form edt_Language_Modify 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Edit Language Record..."
   ClientHeight    =   6000
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5955
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
   Icon            =   "edt_Language_Modify.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   5955
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   WhatsThisHelp   =   -1  'True
   Begin VB.TextBox txtTranslation 
      Height          =   1935
      Left            =   1800
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   2280
      Width           =   3975
   End
   Begin VB.TextBox txtKey 
      Height          =   1935
      Left            =   1800
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   90
      Width           =   3975
   End
   Begin VB.CommandButton btnCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3120
      TabIndex        =   3
      Top             =   5400
      Width           =   1215
   End
   Begin VB.CommandButton btnOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   4440
      TabIndex        =   2
      Top             =   5400
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Press Ctrl+Enter for new line."
      ForeColor       =   &H006A4300&
      Height          =   195
      Left            =   1800
      TabIndex        =   6
      Top             =   4320
      Width           =   4020
   End
   Begin VB.Label Note 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00404040&
      Height          =   615
      Left            =   1800
      TabIndex        =   7
      Top             =   4590
      Width           =   3975
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Translation Value:"
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   0
      TabIndex        =   5
      Top             =   2400
      Width           =   1740
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Translation Key:"
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   0
      TabIndex        =   4
      Top             =   240
      Width           =   1740
   End
End
Attribute VB_Name = "edt_Language_Modify"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim IsBounded As Boolean
Dim KeyBuffer As String
Dim TrBuffer As String

Dim acceModify As Boolean

Dim p As Object
Dim ifEdit As Boolean

Private Sub Form_Load()
    IsBounded = True
End Sub

Private Sub txtKey_GotFocus()
    If ifEdit Then
        If KeyBuffer = "" Then
            Call txtKey.SetFocus
            txtKey.SelStart = 0
            txtKey.SelLength = Len(txtKey.Text)
        Else
            Call txtTranslation.SetFocus
            txtTranslation.SelStart = 0
            txtTranslation.SelLength = Len(txtTranslation.Text)
        End If
    End If
End Sub
Public Sub SetEdit()
    ifEdit = True
End Sub

Friend Sub SetParent(Parent As Object)
    Set p = Parent
End Sub

Private Sub txtKey_Change()
    KeyBuffer = txtKey.Text
    If IsBounded Then
        txtTranslation.Text = KeyBuffer
    End If
    If KeyBuffer <> "" Then _
        acceModify = True
End Sub

Private Sub txtTranslation_Change()
    TrBuffer = txtTranslation.Text
    If TrBuffer = "" Then
        IsBounded = True
    Else
        If txtKey.Text <> TrBuffer Then
            IsBounded = False
        End If
    End If
    acceModify = True
End Sub

Private Sub btnCancel_Click()
    Call Unload(Me)
End Sub

Private Sub btnOK_Click()
    If txtKey.Text = "" Then
        Call MsgBox("Please specify a translation key.", vbCritical + vbOKOnly)
        Exit Sub
    End If
    If Not (p Is Nothing) Then Call p.SetEditorVariables(acceModify, KeyBuffer, TrBuffer)
    Call Unload(Me)
End Sub

Public Property Get TranslationText() As String
    TranslationText = TrBuffer
End Property
Public Property Get KeyText() As String
    KeyText = KeyBuffer
End Property
Public Property Let TranslationText(Value As String)
    TrBuffer = Value
    On Error Resume Next
    txtTranslation.Text = Value
End Property
Public Property Let KeyText(Value As String)
    KeyBuffer = Value
    On Error Resume Next
    txtKey.Text = Value
End Property
Public Property Get Modified() As Boolean
    Modified = acceModify
End Property
Public Property Get NoteInfo() As String
    NoteInfo = Note.Caption
End Property
Public Property Let NoteInfo(Value As String)
    Note.Caption = Value
End Property
