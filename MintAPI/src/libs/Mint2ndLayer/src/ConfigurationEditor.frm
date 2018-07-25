VERSION 5.00
Begin VB.Form frmConfigurationEditor 
   Caption         =   "Configuration Editor"
   ClientHeight    =   11310
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   12945
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
   Icon            =   "ConfigurationEditor.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   11310
   ScaleWidth      =   12945
   StartUpPosition =   2  'CenterScreen
   Begin MintAPI2ndLayer.baEditor Editor 
      Height          =   11175
      Left            =   1710
      TabIndex        =   0
      Top             =   45
      Width           =   11175
      _ExtentX        =   19791
      _ExtentY        =   19711
   End
   Begin VB.Menu m_file 
      Caption         =   "&File"
      Begin VB.Menu m_file_new 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu m_file_open 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu m_file_save 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu m_file_saveas 
         Caption         =   "&Save As"
         Shortcut        =   ^D
      End
      Begin VB.Menu m_file_sep1 
         Caption         =   "-"
      End
      Begin VB.Menu m_file_recent 
         Caption         =   "&Recent Configuration Files"
         Begin VB.Menu m_file_recent_arr 
            Caption         =   "&Item 01"
            Index           =   0
         End
      End
      Begin VB.Menu m_file_sep2 
         Caption         =   "-"
      End
      Begin VB.Menu m_file_import 
         Caption         =   "&Import"
         Shortcut        =   ^I
      End
      Begin VB.Menu m_file_export 
         Caption         =   "&Export"
         Shortcut        =   ^E
      End
      Begin VB.Menu m_file_sep3 
         Caption         =   "-"
      End
      Begin VB.Menu m_file_exit 
         Caption         =   "&Exit"
         Shortcut        =   ^Q
      End
   End
End
Attribute VB_Name = "frmConfigurationEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim p As ConfigurationEditor

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Cancel = True
    Call Me.Hide
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Call Editor.Move(Editor.Left, Editor.Top, ScaleWidth - Editor.Left - 75, ScaleHeight - Editor.Top - 75)
End Sub

Friend Sub SetParent(Parent As ConfigurationEditor)
    Set p = Parent
End Sub

