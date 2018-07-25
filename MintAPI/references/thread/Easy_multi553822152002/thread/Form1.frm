VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Multithread..."
   ClientHeight    =   675
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2790
   LinkTopic       =   "Form1"
   ScaleHeight     =   675
   ScaleWidth      =   2790
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Stop thread"
      Enabled         =   0   'False
      Height          =   555
      Left            =   1440
      TabIndex        =   1
      Top             =   60
      Width           =   1275
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start thread"
      Height          =   555
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    'Set button status
    Me.Command1.Enabled = False
    Me.Command2.Enabled = True
    
    'make sure the thread keeps running until "Stop thread is pressed"
    stopThread = False
    
    'create the thread
    SHCreateThread AddressOf myThread, ByVal 0&, CTF_INSIST, ByVal 0&
    
End Sub

Private Sub Command2_Click()
    'Set button status
    Me.Command1.Enabled = True
    Me.Command2.Enabled = False
    
    'Make the thread stop
    stopThread = True
    
    
End Sub
