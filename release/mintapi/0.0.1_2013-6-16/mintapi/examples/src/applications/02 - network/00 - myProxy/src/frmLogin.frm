VERSION 5.00
Begin VB.Form frmLogin 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login To Proxy Server"
   ClientHeight    =   3045
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6555
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3045
   ScaleWidth      =   6555
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   2070
      TabIndex        =   3
      Text            =   "Text2"
      Top             =   1125
      Width           =   2850
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2070
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   450
      Width           =   2850
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command1"
      Height          =   420
      Left            =   3645
      TabIndex        =   1
      Top             =   2295
      Width           =   1140
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   420
      Left            =   4950
      TabIndex        =   0
      Top             =   2295
      Width           =   1140
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

