VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmSearch 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8115
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   8115
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picHolder 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   435
      Left            =   0
      ScaleHeight     =   435
      ScaleWidth      =   8115
      TabIndex        =   1
      Top             =   2760
      Width           =   8115
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   345
         Left            =   60
         TabIndex        =   2
         Top             =   60
         Width           =   1005
      End
      Begin VB.Label lblPath 
         AutoSize        =   -1  'True
         Caption         =   "lblPath"
         Height          =   195
         Left            =   1320
         TabIndex        =   3
         Top             =   90
         Width           =   480
      End
   End
   Begin MSComctlLib.ListView lvwFiles 
      Height          =   1185
      Left            =   540
      TabIndex        =   0
      Top             =   360
      Width           =   5925
      _ExtentX        =   10451
      _ExtentY        =   2090
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
End
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents mobjNewSearch   As clsNewSearch
Attribute mobjNewSearch.VB_VarHelpID = -1

Private mblnCancel  As Boolean

Private Sub Form_Load()
    With lvwFiles
        .Left = 0
        .Top = 0
        .ColumnHeaders.Add , "FILE", "File"
        .ColumnHeaders.Add , "READONLY", "Read-Only", 1000, lvwColumnCenter
        .ColumnHeaders.Add , "HIDDEN", "Hidden", 1000, lvwColumnCenter
        .View = lvwReport
        .FullRowSelect = True
        .HideSelection = False
        .LabelEdit = lvwManual
    End With
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    With lvwFiles
        .ColumnHeaders.Item("FILE").Width = lvwFiles.Width - 2500
        .Width = Me.ScaleWidth
        .Height = Me.ScaleHeight - picHolder.Height
    End With
End Sub

Friend Property Set NewSearch(ByRef Value As clsNewSearch)
    Set mobjNewSearch = Value
End Property

Friend Sub StartSearch(ByVal pstrPath As String, ByVal pstrExtension As String)
On Error GoTo ErrHandler
    Me.Caption = "File Search: " & pstrPath & "*." & pstrExtension
    FindMultipleFiles pstrPath, pstrExtension
    Me.Caption = "File Search: " & pstrPath & "*." & pstrExtension & " - Completed"
    Exit Sub
ErrHandler:
    If Not (Err.Number = vbObjectError) Then
        Err.Raise Err.Number, Err.Source, Err.Description
    End If
End Sub

Private Sub FindMultipleFiles(ByVal RootPath As String, ByVal FileExtension As String)
Dim Dirs() As String, FoundDirs As Boolean, K As Long
Dim currFile As String
Dim blnReadOnly     As Boolean
Dim blnhidden       As Boolean
Dim lvwItem         As ListItem
On Error GoTo ErrHandler
    If Right(RootPath, 1) <> "\" Then RootPath = RootPath & "\"
    currFile = Dir(RootPath & "*", vbArchive + vbDirectory + vbHidden + vbReadOnly)
    
    Do Until currFile = ""
        DoEvents
        If mblnCancel Then
            Err.Raise vbObjectError
        End If
        If currFile <> "." And currFile <> ".." Then
On Error GoTo MissFile
            lblPath.Caption = RootPath
            If (GetAttr(RootPath & currFile) And vbDirectory) = vbDirectory Then
                If Not FoundDirs Then
                    ReDim Dirs(0)
                Else
                    ReDim Preserve Dirs(UBound(Dirs) + 1)
                End If
                
                Dirs(UBound(Dirs)) = currFile
                FoundDirs = True
            ElseIf Not (InStr(1, currFile, "." & FileExtension) = 0) Or FileExtension = vbNullString Then
                blnReadOnly = (GetAttr(RootPath & currFile) And vbReadOnly) = vbReadOnly
                blnhidden = (GetAttr(RootPath & currFile) And vbHidden) = vbHidden
                Set lvwItem = lvwFiles.ListItems.Add(, , RootPath & currFile)
                lvwItem.SubItems(1) = blnReadOnly
                lvwItem.SubItems(2) = blnhidden
                'lvwItem.EnsureVisible
                Set lvwItem = Nothing
                mobjNewSearch.FoundFile RootPath & currFile, blnReadOnly, blnhidden
            End If
MissFile:
        End If
        
        currFile = Dir
    Loop
    
    If FoundDirs Then
        For K = 0 To UBound(Dirs)
            FindMultipleFiles RootPath & Dirs(K), FileExtension
        Next K
    End If
    Exit Sub
ErrHandler:
    If Err.Number = vbObjectError Then
        Err.Raise Err.Number
    End If
End Sub

Private Sub cmdCancel_Click()
    mblnCancel = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mblnCancel = True
    Set mobjNewSearch = Nothing
End Sub

Private Sub mobjNewSearch_Destroy()
    Set mobjNewSearch = Nothing
    Unload Me
End Sub
