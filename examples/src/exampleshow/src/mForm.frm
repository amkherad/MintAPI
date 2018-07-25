VERSION 5.00
Begin VB.Form mForm 
   BackColor       =   &H00FFFFFF&
   Caption         =   "MintAPI Examples"
   ClientHeight    =   7335
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9900
   BeginProperty Font 
      Name            =   "Consolas"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "mForm.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   489
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   660
   StartUpPosition =   2  'CenterScreen
   Begin examples.uMultiShow uMultiShow1 
      Height          =   6900
      Left            =   45
      TabIndex        =   1
      Top             =   450
      Width           =   9825
      _ExtentX        =   17330
      _ExtentY        =   12171
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   330
      Left            =   45
      TabIndex        =   0
      Top             =   45
      Width           =   1230
   End
End
Attribute VB_Name = "mForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
