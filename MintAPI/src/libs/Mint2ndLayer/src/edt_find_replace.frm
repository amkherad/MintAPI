VERSION 5.00
Begin VB.Form edt_find_replace 
   Caption         =   "Find And Replace..."
   ClientHeight    =   3675
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7500
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
   Icon            =   "edt_find_replace.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3675
   ScaleWidth      =   7500
   StartUpPosition =   1  'CenterOwner
   Begin VB.OptionButton Option2 
      Caption         =   "&Regular Expression"
      Height          =   285
      Left            =   1710
      TabIndex        =   12
      Top             =   3195
      Width           =   2130
   End
   Begin VB.OptionButton Option1 
      Caption         =   "&Extended (\n \r \t \0...)"
      Height          =   285
      Left            =   1710
      TabIndex        =   11
      Top             =   2790
      Width           =   2625
   End
   Begin VB.OptionButton optNormal 
      Caption         =   "&Normal"
      Height          =   285
      Left            =   1710
      TabIndex        =   10
      Top             =   2430
      Width           =   1050
   End
   Begin VB.CommandButton btnReplaceAll 
      Caption         =   "Replace &All"
      Height          =   375
      Left            =   6120
      TabIndex        =   5
      Top             =   1755
      Width           =   1215
   End
   Begin VB.CommandButton btnReplace 
      Caption         =   "&Replace"
      Height          =   375
      Left            =   6120
      TabIndex        =   4
      Top             =   1260
      Width           =   1215
   End
   Begin VB.CheckBox chkMatch 
      Caption         =   "&Match Case"
      Height          =   285
      Left            =   4545
      TabIndex        =   2
      Top             =   2790
      Width           =   1365
   End
   Begin VB.TextBox txtReplace 
      Height          =   700
      Left            =   1710
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   1575
      Width           =   4200
   End
   Begin VB.TextBox txtFindWhat 
      Height          =   700
      Left            =   1710
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   765
      Width           =   4200
   End
   Begin VB.CommandButton btnFind 
      Caption         =   "&Find Next"
      Height          =   375
      Left            =   6120
      TabIndex        =   3
      Top             =   765
      Width           =   1215
   End
   Begin VB.CommandButton btnCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   6120
      TabIndex        =   6
      Top             =   2250
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Replace With:"
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   135
      TabIndex        =   9
      Top             =   1620
      Width           =   1530
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Find What:"
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   135
      TabIndex        =   8
      Top             =   810
      Width           =   1530
   End
   Begin VB.Label lblCaption 
      BackStyle       =   0  'Transparent
      Caption         =   "Find and replace..."
      ForeColor       =   &H006A4300&
      Height          =   510
      Left            =   180
      TabIndex        =   7
      Top             =   135
      Width           =   7170
   End
End
Attribute VB_Name = "edt_find_replace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Enum ButtonType
    btFind
    btReplace
    btReplaceAll
End Enum
Public Enum SearchMode
    smNormal
    smExtended
    smRegularExpression
End Enum
Dim p As Object

Private Sub btnFind_Click()
    If Not p Is Nothing Then
        Call p.FindReplace(ButtonType.btFind, txtFindWhat.Text, txtReplace.Text, GetSearchMode, chkMatch.Value = CheckBoxConstants.vbChecked)
    End If
End Sub
Private Function GetSearchMode() As SearchMode
    
End Function

Public Sub SetParent(Parent As Object)
    Set p = Parent
End Sub

Public Property Get NoteInfo() As String
    NoteInfo = lblCaption.Caption
End Property
Public Property Let NoteInfo(Value As String)
    lblCaption.Caption = Value
End Property

