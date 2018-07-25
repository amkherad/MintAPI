VERSION 5.00
Begin VB.Form frmWidgetDesigner 
   BackColor       =   &H00FFFFFF&
   Caption         =   "MintIntellicense Widget Designer"
   ClientHeight    =   8115
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10770
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   541
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   718
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frmToolbox 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Toolbox"
      Height          =   8055
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   1095
      Begin VB.CommandButton Command8 
         Height          =   375
         Left            =   600
         TabIndex        =   9
         Top             =   1800
         Width           =   375
      End
      Begin VB.CommandButton Command7 
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   1800
         Width           =   375
      End
      Begin VB.CommandButton Command6 
         Height          =   375
         Left            =   600
         TabIndex        =   7
         Top             =   1320
         Width           =   375
      End
      Begin VB.CommandButton Command5 
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   1320
         Width           =   375
      End
      Begin VB.CommandButton Command4 
         Height          =   375
         Left            =   600
         TabIndex        =   5
         Top             =   840
         Width           =   375
      End
      Begin VB.CommandButton Command3 
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   840
         Width           =   375
      End
      Begin VB.CommandButton Command2 
         Height          =   375
         Left            =   600
         TabIndex        =   3
         Top             =   360
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   375
      End
   End
   Begin VB.Frame frmDR 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   8055
      Left            =   1080
      TabIndex        =   0
      Top             =   0
      Width           =   9735
   End
End
Attribute VB_Name = "frmWidgetDesigner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Resize()
    On Error Resume Next
    Call frmToolbox.Move(0, 0, frmToolbox.Width, ScaleHeight)
    Call frmDR.Move(frmToolbox.Width, 0, ScaleWidth - frmToolbox.Width, ScaleHeight)
End Sub
