VERSION 5.00
Begin VB.UserControl FilterUC 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3150
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12660
   BeginProperty Font 
      Name            =   "Consolas"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   3150
   ScaleWidth      =   12660
   Begin VB.Frame frm 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3090
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   12165
      Begin miccyUltimateTools.FilterRow fr 
         Height          =   1290
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   12090
         _ExtentX        =   21325
         _ExtentY        =   2275
      End
   End
   Begin VB.VScrollBar v 
      Height          =   3165
      Left            =   12450
      TabIndex        =   0
      Top             =   0
      Width           =   240
   End
End
Attribute VB_Name = "FilterUC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

