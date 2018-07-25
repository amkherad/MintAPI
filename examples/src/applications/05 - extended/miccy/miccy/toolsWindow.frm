VERSION 5.00
Begin VB.Form toolsWindow 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "miccy Picture Tool version 1.0.1"
   ClientHeight    =   7275
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13395
   BeginProperty Font 
      Name            =   "Consolas"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "toolsWindow.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7275
   ScaleWidth      =   13395
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton saveButton 
      Caption         =   "&Save"
      Height          =   390
      Left            =   10575
      TabIndex        =   0
      Top             =   5550
      Width           =   1365
   End
End
Attribute VB_Name = "toolsWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

