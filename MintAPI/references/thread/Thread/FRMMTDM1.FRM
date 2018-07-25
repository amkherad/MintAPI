VERSION 5.00
Begin VB.Form frmMTDemo1 
   Caption         =   "Form1"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2790
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   2790
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Launch New Form"
      Height          =   375
      Left            =   360
      TabIndex        =   8
      Top             =   2700
      Width           =   1995
   End
   Begin VB.Timer Timer1 
      Left            =   2340
      Top             =   180
   End
   Begin VB.Label lblVerification 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   540
      TabIndex        =   7
      Top             =   2340
      Width           =   1635
   End
   Begin VB.Label lblOperation 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   540
      TabIndex        =   6
      Top             =   1710
      Width           =   2175
   End
   Begin VB.Label lblAccumulator 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   540
      TabIndex        =   5
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label lblGlobalCounter 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   540
      TabIndex        =   4
      Top             =   450
      Width           =   1185
   End
   Begin VB.Label Label3 
      Caption         =   "TotalIncrements:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   180
      TabIndex        =   3
      Top             =   1980
      Width           =   1725
   End
   Begin VB.Label Label2 
      Caption         =   "Operation:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   180
      TabIndex        =   2
      Top             =   1440
      Width           =   1905
   End
   Begin VB.Label Label1 
      Caption         =   "Accumulator:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   180
      TabIndex        =   1
      Top             =   810
      Width           =   1815
   End
   Begin VB.Label label4 
      Caption         =   "GenericGlobalCounter:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   180
      TabIndex        =   0
      Top             =   180
      Width           =   1995
   End
End
Attribute VB_Name = "frmMTDemo1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' MTDemo - Multithreading Demo program
' Copyright © 1997 by Desaware Inc. All Rights Reserved
Option Explicit

Dim State As Integer
' State = 0 - Idle
' State = 1 - Loading existing value
' State = 2 - Adding 1 to existing value
' State = 3 - Storing existing value
' State = 4 - Extra delay

Dim Accumulator As Long

Const OtherCodeDelay = 10

Private Sub Command1_Click()
   Dim f As New frmMTDemo1
   f.Show
End Sub

Private Sub Form_Load()
   Timer1.Interval = 750 + Rnd * 500
End Sub

Private Sub Timer1_Timer()
   Static otherdelay&
   Select Case State
      Case 0
         lblOperation = "Idle"
         State = 1
      Case 1
         lblOperation = "Loading Acc"
         Accumulator = GenericGlobalCounter
         State = 2
      Case 2
         lblOperation = "Incrementing"
         Accumulator = Accumulator + 1
         State = 3
      Case 3
         lblOperation = "Storing"
         GenericGlobalCounter = Accumulator
         TotalIncrements = TotalIncrements + 1
         State = 4
      Case 4
         lblOperation = "Generic Code"
         If otherdelay >= OtherCodeDelay Then
            State = 0
            otherdelay = 0
         Else
            otherdelay = otherdelay + 1
         End If
   End Select
   UpdateDisplay
   
End Sub

Public Sub UpdateDisplay()
   lblGlobalCounter = Str$(GenericGlobalCounter)
   lblAccumulator = Str$(Accumulator)
   lblVerification = Str$(TotalIncrements)
End Sub
