VERSION 5.00
Begin VB.Form LanguageEditor 
   Caption         =   "MintAPI Language Editor"
   ClientHeight    =   6825
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   12600
   BeginProperty Font 
      Name            =   "Consolas"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "LanguageEditor.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6825
   ScaleWidth      =   12600
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame pnl 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   390
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Visible         =   0   'False
      Width           =   915
      Begin VB.CommandButton btn_add_Cancel 
         Caption         =   "&Cancel"
         Height          =   390
         Left            =   4500
         TabIndex        =   18
         Top             =   3375
         Width           =   1215
      End
      Begin VB.CommandButton btn_add_OK 
         Caption         =   "&OK"
         Height          =   390
         Left            =   5775
         TabIndex        =   17
         Top             =   3375
         Width           =   1215
      End
      Begin VB.CheckBox add_rtl 
         Caption         =   "Right To Left Display"
         Height          =   240
         Left            =   2400
         TabIndex        =   16
         Top             =   2775
         Width           =   2190
      End
      Begin VB.TextBox add_desc 
         Height          =   840
         Left            =   2400
         MaxLength       =   100
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   15
         Top             =   1800
         Width           =   4590
      End
      Begin VB.TextBox add_region 
         Height          =   315
         Left            =   2400
         MaxLength       =   20
         TabIndex        =   13
         Top             =   1350
         Width           =   1890
      End
      Begin VB.TextBox add_shortname 
         Height          =   315
         Left            =   2400
         MaxLength       =   2
         TabIndex        =   11
         Text            =   "en"
         Top             =   900
         Width           =   315
      End
      Begin VB.TextBox add_name 
         Height          =   315
         Left            =   2400
         MaxLength       =   50
         TabIndex        =   9
         Top             =   450
         Width           =   4590
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Description:"
         Height          =   195
         Left            =   600
         TabIndex        =   14
         Top             =   1860
         Width           =   1740
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Region:"
         Height          =   195
         Left            =   600
         TabIndex        =   12
         Top             =   1410
         Width           =   1740
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Short Name:"
         Height          =   195
         Left            =   600
         TabIndex        =   10
         Top             =   960
         Width           =   1740
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Name:"
         Height          =   195
         Left            =   600
         TabIndex        =   8
         Top             =   505
         Width           =   1740
      End
   End
   Begin VB.Frame frmlng 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   5340
      Left            =   0
      TabIndex        =   19
      Top             =   1500
      Width           =   12615
      Begin VB.VScrollBar scr 
         Height          =   4965
         Left            =   12375
         TabIndex        =   21
         Top             =   0
         Width           =   240
      End
      Begin VB.Frame frm 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         ForeColor       =   &H80000008&
         Height          =   4665
         Left            =   0
         TabIndex        =   20
         Top             =   0
         Width           =   12165
      End
   End
   Begin VB.CommandButton removeLanguage 
      Caption         =   "&Delete"
      Height          =   315
      Left            =   8700
      TabIndex        =   6
      Top             =   600
      Width           =   915
   End
   Begin VB.CommandButton addnewLanguage 
      Caption         =   "+"
      Height          =   315
      Left            =   8175
      TabIndex        =   5
      Top             =   600
      Width           =   465
   End
   Begin VB.ComboBox cfile 
      Height          =   315
      Left            =   2700
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   600
      Width           =   5415
   End
   Begin VB.CommandButton wBrowse 
      Caption         =   "..."
      Height          =   315
      Left            =   8175
      TabIndex        =   2
      Top             =   150
      Width           =   465
   End
   Begin VB.TextBox wPath 
      Height          =   315
      Left            =   2700
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   150
      Width           =   5415
   End
   Begin VB.Line lnSep 
      X1              =   4200
      X2              =   4200
      Y1              =   1275
      Y2              =   1500
   End
   Begin VB.Line ln 
      X1              =   0
      X2              =   12600
      Y1              =   1275
      Y2              =   1275
   End
   Begin VB.Label keyLBL 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Language Key"
      Height          =   240
      Left            =   0
      TabIndex        =   22
      Top             =   1275
      Width           =   4140
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Current Translation File:"
      Height          =   195
      Left            =   225
      TabIndex        =   4
      Top             =   660
      Width           =   2415
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Working Directory:"
      Height          =   195
      Left            =   900
      TabIndex        =   0
      Top             =   205
      Width           =   1740
   End
   Begin VB.Label translationLBL 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Language Translation Value"
      Height          =   240
      Left            =   4125
      TabIndex        =   23
      Top             =   1275
      Width           =   8415
   End
   Begin VB.Menu m_file 
      Caption         =   "&File"
      Begin VB.Menu m_file_new 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu m_file_open 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu m_file_save 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu m_file_saveas 
         Caption         =   "Save &As"
         Shortcut        =   ^E
      End
      Begin VB.Menu m_file_sep1 
         Caption         =   "-"
      End
      Begin VB.Menu m_file_import 
         Caption         =   "&Import"
      End
      Begin VB.Menu m_file_export 
         Caption         =   "&Export"
      End
      Begin VB.Menu m_file_sep2 
         Caption         =   "-"
      End
      Begin VB.Menu m_file_compile 
         Caption         =   "&Compile Language File"
         Shortcut        =   ^H
      End
      Begin VB.Menu m_file_sep3 
         Caption         =   "-"
      End
      Begin VB.Menu m_file_exit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu m_edit 
      Caption         =   "&Edit"
      Begin VB.Menu m_edit_undo 
         Caption         =   "&Undo"
         Shortcut        =   ^Z
      End
      Begin VB.Menu m_edit_redo 
         Caption         =   "&Redo"
         Shortcut        =   ^Y
      End
      Begin VB.Menu m_edit_sep1 
         Caption         =   "-"
      End
      Begin VB.Menu m_edit_cut 
         Caption         =   "&Cut"
         Shortcut        =   ^X
      End
      Begin VB.Menu m_edit_copy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu m_edit_paste 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu m_edit_selectall 
         Caption         =   "&Select All"
         Shortcut        =   ^A
      End
      Begin VB.Menu m_edit_sep2 
         Caption         =   "-"
      End
      Begin VB.Menu m_edit_find 
         Caption         =   "&Find"
         Shortcut        =   ^F
      End
      Begin VB.Menu m_edit_replace 
         Caption         =   "&Replace"
         Shortcut        =   ^R
      End
      Begin VB.Menu m_edit_sep3 
         Caption         =   "-"
      End
      Begin VB.Menu m_edit_duplicate 
         Caption         =   "&Duplicate Selected Record"
         Shortcut        =   ^D
      End
      Begin VB.Menu m_edit_append 
         Caption         =   "&Append Record"
      End
      Begin VB.Menu m_edit_delete 
         Caption         =   "Delete &Selected Record"
      End
      Begin VB.Menu m_edit_sep4 
         Caption         =   "-"
      End
      Begin VB.Menu m_edit_advanced 
         Caption         =   "&Advanced"
         Begin VB.Menu m_edit_adv_batchtranslation 
            Caption         =   "&Batch Translation"
         End
      End
      Begin VB.Menu m_edit_sep5 
         Caption         =   "-"
      End
      Begin VB.Menu m_edit_translationsettings 
         Caption         =   "&Translation File Settings"
         Shortcut        =   ^B
      End
   End
   Begin VB.Menu m_about 
      Caption         =   "&About"
      Begin VB.Menu m_about_manual 
         Caption         =   "&Manual"
         Shortcut        =   {F1}
      End
      Begin VB.Menu m_about_aboutlang 
         Caption         =   "&About MintAPI Language Editor"
      End
      Begin VB.Menu m_about_mintapi 
         Caption         =   "&About MintAPI"
         Shortcut        =   {F2}
      End
   End
End
Attribute VB_Name = "LanguageEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim columnKeys As Long
Dim columnValues As Long

Private Sub addnewLanguage_Click()
    If Directory(wPath.Text).Exists Then
        pnl.Visible = True
        Call pnl.Move(0, 0, ScaleWidth, ScaleHeight)
    Else
        Call MsgBox("Please select language directory first ,where created file will be placed.", vbCritical, "Invalid Directory Path")
    End If
End Sub

Private Sub Form_Load()
On Error Resume Next
    
End Sub

Public Sub SetLangPath(ByVal Path As String)
     wPath.Text = Path
End Sub

Private Sub removeLanguage_Click()
    '
End Sub

Private Sub btn_add_Cancel_Click()
    pnl.Visible = False
End Sub
Private Sub btn_add_OK_Click()
    pnl.Visible = False
End Sub

Private Sub Form_Resize()
On Error Resume Next
    If pnl.Visible Then
        Call pnl.Move(0, 0, ScaleWidth, ScaleHeight)
    End If
    Call frmlng.Move(0, frmlng.Top, ScaleWidth, ScaleHeight - frmlng.Top)
    ln.X2 = Width
    columnKeys = (ScaleWidth * 30) / 100
    columnValues = (ScaleWidth * 70) / 100
    lnSep.X1 = columnKeys
    lnSep.X2 = columnKeys
    keyLBL.Width = columnKeys
    translationLBL.Left = columnKeys + 1
    translationLBL.Width = columnValues - 1
    scr.Left = ScaleWidth - scr.Width
    scr.Height = frmlng.Height
    frm.Width = scr.Left
End Sub

Private Sub wBrowse_Click()
On Error GoTo canceled
'"All Language Files(*.lng;*.lang;*.mlng;*.mintapilanguage)|*.lng;*.lang;*.mlng;*.mintapilanguage|Language Files(*.lng;*.lang)|*.lng;*.lang|MintAPI Language Files(*.mlng;*.mintapilanguage)|*.mlng;*.mintapilanguage|All Files(*)|*"
    wPath.Text = Directory.ChooseDirectory(hWnd, "Language Files Directory...", sfCUSTOM, , IIf(wPath.Text = "", Directory.CurrentDirectory, Directory(wPath.Text)), True, False, False).AbsolutePath
    'wPath.Text = File.ChooseFile(fdSave, "Language File...", "All Files (*.*)" & vbCrLf & "*.*", Nothing, Me.hWnd).Directory.AbsolutePath
    Dim f() As String
    f = Directory(wPath.Text).FilteredFileNames(FileFilters(StringArray("*.mts", "*.lang", "*.lng", "*.mlang", "*.mintapilanguage"), StringArray("")))
    Dim i As Long, Count As Long
    Count = ArraySize(f)
    Call cfile.Clear
    For i = 0 To Count - 1
        Call cfile.AddItem(f(i))
    Next
    If Count > 0 Then
        cfile.ListIndex = 0
    End If
canceled:
End Sub
