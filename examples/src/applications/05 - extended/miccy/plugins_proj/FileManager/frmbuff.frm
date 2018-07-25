VERSION 5.00
Begin VB.Form frmbuff 
   ClientHeight    =   7755
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15030
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
   LinkTopic       =   "Form1"
   ScaleHeight     =   7755
   ScaleWidth      =   15030
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame masterProps 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H00404040&
      Height          =   3540
      Left            =   0
      TabIndex        =   1
      Top             =   300
      Width           =   13740
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Include Sub Folders In Proccess"
         ForeColor       =   &H00404040&
         Height          =   240
         Left            =   450
         TabIndex        =   7
         Top             =   525
         Width           =   3090
      End
      Begin VB.CheckBox incSubFolders 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Include Sub Folders In Proccess"
         ForeColor       =   &H00404040&
         Height          =   240
         Left            =   450
         TabIndex        =   6
         Top             =   2700
         Width           =   3090
      End
      Begin VB.TextBox Text1 
         ForeColor       =   &H00404040&
         Height          =   1440
         Left            =   7875
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   1005
         Width           =   4065
      End
      Begin VB.TextBox incFormats 
         ForeColor       =   &H00404040&
         Height          =   1440
         Left            =   1950
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   1005
         Width           =   4065
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Exclude Formats:"
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   6375
         TabIndex        =   4
         Top             =   1050
         Width           =   1440
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Include Formats:"
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   450
         TabIndex        =   2
         Top             =   1050
         Width           =   1440
      End
   End
   Begin VB.Frame FileManager 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H00404040&
      Height          =   6465
      Left            =   975
      TabIndex        =   0
      Top             =   450
      Width           =   12240
   End
End
Attribute VB_Name = "frmbuff"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

