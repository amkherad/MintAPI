VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Waiting"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6600
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   6600
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkErrors 
      Caption         =   "Handle Errors"
      Height          =   285
      Left            =   840
      TabIndex        =   13
      Top             =   1740
      Width           =   1815
   End
   Begin VB.CommandButton cmdError 
      Caption         =   "Raise Error"
      Height          =   435
      Left            =   870
      TabIndex        =   12
      Top             =   2580
      Width           =   1365
   End
   Begin VB.CommandButton cmdPostData 
      Caption         =   "Post Data"
      Height          =   435
      Left            =   2340
      TabIndex        =   11
      Top             =   2580
      Width           =   1365
   End
   Begin VB.ListBox lstData 
      Height          =   2790
      Left            =   4230
      TabIndex        =   10
      Top             =   150
      Width           =   2235
   End
   Begin VB.TextBox txthWnd 
      BackColor       =   &H8000000B&
      Height          =   345
      Left            =   1470
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   150
      Width           =   2625
   End
   Begin VB.CommandButton cmdPing 
      Caption         =   "Ping"
      Height          =   435
      Left            =   870
      TabIndex        =   7
      Top             =   2100
      Width           =   1365
   End
   Begin VB.CommandButton cmdSendData 
      Caption         =   "Send Data"
      Height          =   435
      Left            =   2340
      TabIndex        =   6
      Top             =   2100
      Width           =   1365
   End
   Begin VB.TextBox txtChildhWnd 
      Height          =   345
      Left            =   1470
      TabIndex        =   4
      Top             =   540
      Width           =   2625
   End
   Begin VB.TextBox txtReplyData 
      Height          =   345
      Left            =   1470
      TabIndex        =   2
      Text            =   "Sausage"
      Top             =   1320
      Width           =   2625
   End
   Begin VB.TextBox txtSendData 
      Height          =   345
      Left            =   1470
      TabIndex        =   0
      Text            =   "Growl"
      Top             =   930
      Width           =   2625
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "hWnd:"
      Height          =   255
      Left            =   390
      TabIndex        =   9
      Top             =   210
      Width           =   1035
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Child hWnd:"
      Height          =   255
      Left            =   390
      TabIndex        =   5
      Top             =   600
      Width           =   1035
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Reply Data:"
      Height          =   255
      Left            =   390
      TabIndex        =   3
      Top             =   1380
      Width           =   1035
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Send Data:"
      Height          =   255
      Left            =   390
      TabIndex        =   1
      Top             =   990
      Width           =   1035
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents mobjGateway As Gateway
Attribute mobjGateway.VB_VarHelpID = -1
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Sub chkErrors_Click()
    mobjGateway.HandleAsyncErrors = (chkErrors.Value = vbChecked)
End Sub

Private Sub cmdError_Click()
On Error GoTo ErrHandler
    MsgBox 1 / 0
    Exit Sub
ErrHandler:
    mobjGateway.RaiseError Err.Number, Err.Source, Err.Description
End Sub

Private Sub cmdPing_Click()
    mobjGateway.Ping
End Sub

Private Sub cmdPostData_Click()
Dim bytData()   As Byte
On Error GoTo ErrHandler
    bytData = txtSendData.Text
    mobjGateway.PostData bytData
    Exit Sub
ErrHandler:
    MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "ERROR"
End Sub

Private Sub cmdSendData_Click()
Dim strReturn   As String
Dim bytData()   As Byte
On Error GoTo ErrHandler
    bytData = txtSendData.Text
    strReturn = mobjGateway.SendData(bytData)
    lstData.AddItem "(R) - " & strReturn
    Exit Sub
ErrHandler:
    MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "ERROR"
End Sub

Private Sub Form_Load()
    Set mobjGateway = New Gateway
    txthWnd.Text = mobjGateway.hWnd
End Sub


Private Sub mobjGateway_AsynchronousError(ByVal Number As Long, ByVal Source As String, ByVal Description As String)
Dim strMsg  As String
    strMsg = "This is an error that has been raises asynchronously."
    strMsg = strMsg & vbCrLf & vbCrLf
    strMsg = strMsg & Number & vbCrLf
    strMsg = strMsg & Source & vbCrLf
    strMsg = strMsg & Description & vbCrLf
    MsgBox strMsg, vbCritical, "Async Error Handler"
End Sub

Private Sub mobjGateway_DataArrived(Data As Variant, ByVal Synchronous As Boolean)
Dim bytReturn()     As Byte
On Error GoTo ErrHandler
    lstData.AddItem "(A) - " & Data
    If CStr(Data) = "ERROR" Then
        Err.Raise 23, "Woof", "This is a forced error!"
    End If
    bytReturn = txtReplyData.Text
    mobjGateway.Reply = bytReturn
    lstData.Refresh
    If Synchronous Then
        Sleep 500
    End If
    Exit Sub
ErrHandler:
    mobjGateway.RaiseError Err.Number, Err.Source, Err.Description
End Sub

Private Sub mobjGateway_LinkInitialised()
    Me.Caption = "Connected"
End Sub

Private Sub mobjGateway_LinkTerminated()
    Me.Caption = "Disconnected"
End Sub

Private Sub txtChildhWnd_Change()
On Error GoTo ErrHandler
    mobjGateway.StartLink Val(txtChildhWnd.Text)
    Exit Sub
ErrHandler:
    MsgBox Err.Number & vbCrLf & Err.Description
End Sub
