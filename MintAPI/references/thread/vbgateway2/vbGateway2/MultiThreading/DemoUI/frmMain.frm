VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Multithreading File Search"
   ClientHeight    =   4560
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9045
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4560
   ScaleWidth      =   9045
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdNew 
      Caption         =   "New Search"
      Height          =   315
      Left            =   2700
      TabIndex        =   3
      Top             =   540
      Width           =   1425
   End
   Begin VB.TextBox txtPath 
      Height          =   315
      Left            =   1140
      TabIndex        =   2
      Text            =   "c:\"
      Top             =   120
      Width           =   2985
   End
   Begin VB.TextBox txtExt 
      Height          =   315
      Left            =   1140
      TabIndex        =   0
      Top             =   510
      Width           =   855
   End
   Begin MSComctlLib.ListView lvwSearch 
      Height          =   3435
      Left            =   90
      TabIndex        =   1
      Top             =   1020
      Width           =   8805
      _ExtentX        =   15531
      _ExtentY        =   6059
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Path:"
      Height          =   225
      Left            =   90
      TabIndex        =   5
      Top             =   150
      Width           =   1005
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Extension:"
      Height          =   225
      Left            =   90
      TabIndex        =   4
      Top             =   540
      Width           =   1005
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents mobjSearches     As clsSearches
Attribute mobjSearches.VB_VarHelpID = -1

Private Sub cmdNew_Click()
    mobjSearches.CreateNew txtPath.Text, txtExt.Text
End Sub

Private Sub Form_Load()
    With lvwSearch
        .ColumnHeaders.Add , "HWND", "hWnd", 1000
        .ColumnHeaders.Add , "COUNT", "File Count", 1000, lvwColumnCenter
        .ColumnHeaders.Add , "CURRENT", "Current File", 9000
        .View = lvwReport
        .FullRowSelect = True
        .HideSelection = False
        .LabelEdit = lvwManual
    End With
    Set mobjSearches = New clsSearches
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mobjSearches = Nothing
End Sub

Private Sub mobjSearches_FileFound(ByVal pstrKey As String, ByVal plnghWnd As Long, ByVal plngCount As Long, ByVal pstrCurrentFile As String)
Dim lvwItem     As ListItem
    Set lvwItem = lvwSearch.ListItems.Item(pstrKey)
    lvwItem.SubItems(1) = plngCount
    lvwItem.SubItems(2) = pstrCurrentFile
    Set lvwItem = Nothing
End Sub

Private Sub mobjSearches_NewSearch(ByVal pstrKey As String, ByVal plnghWnd As Long)
Dim lvwItem As ListItem
    Set lvwItem = lvwSearch.ListItems.Add(, pstrKey, plnghWnd)
    lvwItem.SubItems(1) = "0"
    lvwItem.SubItems(2) = "N/A"
    Set lvwItem = Nothing
End Sub

Private Sub mobjSearches_SearchRemoved(ByVal pstrKey As String)
    lvwSearch.ListItems.Remove pstrKey
End Sub
