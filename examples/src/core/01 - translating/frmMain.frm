VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "MintAPI Translation Test"
   ClientHeight    =   3060
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6210
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3060
   ScaleWidth      =   6210
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Bound"
      Height          =   1860
      Left            =   3195
      TabIndex        =   6
      Top             =   1080
      Width           =   2940
      Begin VB.Label kvb 
         Alignment       =   2  'Center
         Caption         =   "Key = Value"
         Height          =   240
         Left            =   45
         TabIndex        =   12
         Tag             =   "Key = Value"
         Top             =   1260
         Width           =   2850
      End
      Begin VB.Label tb 
         Alignment       =   2  'Center
         Caption         =   "Test"
         Height          =   240
         Left            =   45
         TabIndex        =   11
         Tag             =   "Test"
         Top             =   945
         Width           =   2850
      End
      Begin VB.Label mtb 
         Alignment       =   2  'Center
         Caption         =   "MintAPI Translating"
         Height          =   240
         Left            =   45
         TabIndex        =   10
         Tag             =   "MintAPI Translating"
         Top             =   630
         Width           =   2850
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Unbound"
      Height          =   1860
      Left            =   90
      TabIndex        =   4
      Top             =   1080
      Width           =   2940
      Begin VB.CommandButton refreshRecords 
         Caption         =   "Refresh"
         Height          =   330
         Left            =   1215
         TabIndex        =   5
         Top             =   180
         Width           =   1635
      End
      Begin VB.Label kv 
         Alignment       =   2  'Center
         Caption         =   "Key = Value"
         Height          =   240
         Left            =   45
         TabIndex        =   9
         Tag             =   "Key = Value"
         Top             =   1260
         Width           =   2850
      End
      Begin VB.Label t 
         Alignment       =   2  'Center
         Caption         =   "Test"
         Height          =   240
         Left            =   45
         TabIndex        =   8
         Tag             =   "Test"
         Top             =   945
         Width           =   2850
      End
      Begin VB.Label mt 
         Alignment       =   2  'Center
         Caption         =   "MintAPI Translating"
         Height          =   240
         Left            =   45
         TabIndex        =   7
         Tag             =   "MintAPI Translating"
         Top             =   630
         Width           =   2850
      End
   End
   Begin VB.CommandButton refBtn 
      Caption         =   "Refresh"
      Height          =   330
      Left            =   5265
      TabIndex        =   3
      Top             =   540
      Width           =   870
   End
   Begin VB.ComboBox langs 
      Height          =   315
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   540
      Width           =   4155
   End
   Begin VB.Label Label2 
      Caption         =   "Languages loads from static path .../languages"
      Height          =   555
      Left            =   180
      TabIndex        =   2
      Top             =   90
      Width           =   5415
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Language:"
      Height          =   195
      Left            =   180
      TabIndex        =   1
      Top             =   585
      Width           =   840
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim loadedFiles() As String
Dim cLang As Language

Private Sub LoadLanguages()
    Call langs.Clear
    Dim dir As Directory
    Set dir = CurrentDirectory
    dir.cd "languages"
    If dir.Exists Then
        Dim lngFiles() As String
        lngFiles = dir.FilteredFileNames(Generics.FileFilters(S("*.lng", "*.tr", "*.mlng"), ""), vbNormal, True)
        loadedFiles = lngFiles
        Dim i As Long
        For i = 0 To ArraySize(loadedFiles) - 1
            Call langs.AddItem(File(loadedFiles(i)).NameOnly)
        Next
    Else
        dir.Create
        MsgBox "No languages found in [PATH]/languages"
    End If
    
    If ArraySize(loadedFiles) > 0 Then
        langs.ListIndex = 0
        SetLanguage loadedFiles(0)
    End If
    
    Application.boundtr CStr(mtb.Tag), mtb, "Caption", True
    Application.boundtr CStr(tb.Tag), tb, "Caption", True
    Application.boundtr CStr(kvb.Tag), kvb, "Caption", True
End Sub
Private Sub SetLanguage(Path As String)
    If Not cLang Is Nothing Then _
        Call Application.UnRegisterLanguage(cLang.Name)
    Set cLang = Language(Path)
    MsgBox cLang.Name
    cLang.Load
    
    Application.RegisterLanguage cLang, True
End Sub

Private Sub Form_Load()
    LoadLanguages
End Sub

Private Sub langs_Change()
    SetLanguage loadedFiles(langs.ListIndex)
End Sub

Private Sub langs_Click()
    SetLanguage loadedFiles(langs.ListIndex)
End Sub

Private Sub refBtn_Click()
    LoadLanguages
End Sub

Private Sub refreshRecords_Click()
    mt.Caption = tr(mt.Tag)
    t.Caption = tr(t.Tag)
    kv.Caption = tr(kv.Tag)
End Sub
