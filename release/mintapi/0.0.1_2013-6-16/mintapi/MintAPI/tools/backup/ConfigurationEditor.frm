VERSION 5.00
Begin VB.Form ConfigurationEditor 
   Caption         =   "MintAPI Configuration Editor"
   ClientHeight    =   7080
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13185
   BeginProperty Font 
      Name            =   "Consolas"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "ConfigurationEditor.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7080
   ScaleWidth      =   13185
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton configPathBrowser 
      Caption         =   "..."
      Height          =   300
      Left            =   7275
      TabIndex        =   4
      Top             =   1100
      Width           =   540
   End
   Begin VB.TextBox configPath 
      Height          =   300
      Left            =   2475
      TabIndex        =   3
      Text            =   "[STARTUPPATH]\config.ini"
      Top             =   1100
      Width           =   4740
   End
   Begin VB.OptionButton spcConfig 
      Caption         =   "Specified Application Configuration File"
      Height          =   240
      Left            =   975
      TabIndex        =   1
      Top             =   600
      Width           =   3840
   End
   Begin VB.OptionButton defConfig 
      Caption         =   "Default Application Configuration File"
      Height          =   240
      Left            =   975
      TabIndex        =   0
      Top             =   300
      Value           =   -1  'True
      Width           =   3690
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Configuration File Path:"
      Height          =   195
      Left            =   225
      TabIndex        =   2
      Top             =   1125
      Width           =   2190
   End
End
Attribute VB_Name = "ConfigurationEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public VBInstance As VBIDE.VBE
Public Connect As Connect

Private Sub CancelButton_Click()
    Connect.Hide
End Sub

Private Sub OKButton_Click()
    MsgBox "AddIn operation on: " & VBInstance.FullName
End Sub


