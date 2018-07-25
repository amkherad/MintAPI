VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7590
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6450
   LinkTopic       =   "Form1"
   ScaleHeight     =   7590
   ScaleWidth      =   6450
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Clear Log"
      Height          =   375
      Left            =   240
      TabIndex        =   14
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Caption         =   "Monitor SubDirectory?"
      Height          =   615
      Left            =   2640
      TabIndex        =   11
      Top             =   1680
      Width           =   2895
      Begin VB.OptionButton Option2 
         Caption         =   "No"
         Height          =   255
         Left            =   1560
         TabIndex        =   13
         Top             =   240
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Yes"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Specify Flags"
      Height          =   1335
      Left            =   2640
      TabIndex        =   5
      Top             =   240
      Width           =   2895
      Begin VB.CheckBox CL_Write 
         Caption         =   "Last write"
         Height          =   255
         Left            =   1440
         TabIndex        =   10
         Top             =   600
         Width           =   1095
      End
      Begin VB.CheckBox CL_Access 
         Caption         =   "Last Access"
         Height          =   255
         Left            =   1440
         TabIndex        =   9
         Top             =   240
         Width           =   1335
      End
      Begin VB.CheckBox CF_size 
         Caption         =   "File size"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   600
         Width           =   855
      End
      Begin VB.CheckBox CD_name 
         Caption         =   "Dir name"
         Height          =   195
         Left            =   240
         TabIndex        =   7
         Top             =   960
         Width           =   975
      End
      Begin VB.CheckBox CF_name 
         Caption         =   "File name"
         Height          =   195
         Left            =   240
         TabIndex        =   6
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.TextBox Text1 
      Height          =   4455
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   4
      Top             =   3000
      Width           =   6255
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   3
      Top             =   240
      Width           =   2055
   End
   Begin VB.DirListBox Dir1 
      Height          =   1665
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Start! Monitor"
      Height          =   375
      Left            =   3840
      TabIndex        =   1
      Top             =   2520
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Stop Monitor"
      Height          =   375
      Left            =   5040
      TabIndex        =   0
      Top             =   2520
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'A File System Monitor for VB6 as an example of using modThreading.bas to create threads
'It is in the form of an application, however if you guys want to create a usercontrol out
'of it then do submit in PSC after you've created it.

Option Explicit
Dim M_Path As String, Flags As NotificationFlags
Dim No As Integer
Private Declare Function GetShortPathName Lib "kernel32.dll" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
'Dim WithEvents MonitorThread As Thread
Dim WithEvents Monitor As DirectoryWatcher
Attribute Monitor.VB_VarHelpID = -1

Private Sub Command1_Click()
Command1.Enabled = False
Monitor.StopMonitor
End Sub

Private Sub Command2_Click()
'Set the flags, path and create a monitor and then start it
On Error Resume Next


Flags = DIR_NAME * CD_name.Value Or _
        FILE_NAME * CF_name.Value Or _
        LAST_ACCESS * CL_Access.Value Or _
        LAST_WRITE * CL_Write.Value Or _
        FILESIZE * CF_size.Value

M_Path = Dir1.Path

Monitor.CreateMonitor Me, M_Path, Option1.Value, Flags


Dim st As String * 30

GetShortPathName M_Path, M_Path, Len(M_Path)
M_Path = Left(M_Path, InStr(1, M_Path, vbNullChar) - 1)
ADDtext "Monitor Added and started. Path is" & vbCrLf & M_Path
Command1.Enabled = True

End Sub

Private Sub Command3_Click()
Text1.Text = ""
No = 0
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
End Sub

Public Sub ADDtext(txt As String)
No = No + 1
Text1.Text = Text1.Text & Str(No) & ".)  " & txt & vbCrLf
End Sub

Private Sub Form_Load()
'Check the Version.
Set Monitor = New DirectoryWatcher
ADDtext "App started on ==> " & Now
End Sub

Private Sub Monitor_OnChanged(ByVal Name As String)
ADDtext Monitor.GetFullPath(Name) & " was modified."
End Sub

Private Sub Monitor_OnCreated(ByVal Name As String)
ADDtext Monitor.GetFullPath(Name) & " was created."
End Sub

Private Sub Monitor_OnDeleted(ByVal Name As String)
ADDtext Monitor.GetFullPath(Name) & " was deleted."
End Sub

Private Sub Monitor_OnRenamed(ByVal OldName As String, ByVal NewName As String)
ADDtext Monitor.GetFullPath(OldName) & " was renamed into " & Monitor.GetFullPath(NewName)
End Sub

Private Sub Monitor_OnStopped()
ADDtext "Monitor Stopped."
End Sub
