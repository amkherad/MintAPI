VERSION 5.00
Begin VB.Form ProjectManager 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "MintAPI Project Manager"
   ClientHeight    =   6930
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   12330
   BeginProperty Font 
      Name            =   "Consolas"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "ProjectManager.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6930
   ScaleWidth      =   12330
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox Languaged 
      Caption         =   "Generate Application Language File"
      Height          =   195
      Left            =   2100
      TabIndex        =   21
      Top             =   2475
      Value           =   1  'Checked
      Width           =   3450
   End
   Begin VB.TextBox txtDesc 
      Height          =   285
      Left            =   6675
      TabIndex        =   19
      Text            =   "MintAPI_Project"
      Top             =   300
      Width           =   3900
   End
   Begin VB.CommandButton btnApply 
      Caption         =   "&Apply"
      Height          =   390
      Left            =   8700
      TabIndex        =   18
      Top             =   6075
      Width           =   1365
   End
   Begin VB.OptionButton fileConfig 
      Caption         =   "File Configuration Based"
      Height          =   195
      Left            =   2340
      TabIndex        =   17
      Top             =   1725
      Value           =   -1  'True
      Width           =   2475
   End
   Begin VB.OptionButton regConfig 
      Caption         =   "Registry Configuration Based"
      Height          =   195
      Left            =   2340
      TabIndex        =   16
      Top             =   2025
      Width           =   2850
   End
   Begin VB.ComboBox clang 
      Height          =   315
      ItemData        =   "ProjectManager.frx":2370A
      Left            =   2100
      List            =   "ProjectManager.frx":2371A
      Style           =   2  'Dropdown List
      TabIndex        =   14
      Top             =   3825
      Width           =   3600
   End
   Begin VB.CommandButton btnEditConfig 
      Caption         =   "Edit Configuration"
      Height          =   420
      Left            =   6675
      TabIndex        =   13
      Top             =   1350
      Width           =   2550
   End
   Begin VB.CommandButton btnEditLang 
      Caption         =   "Edit Language File"
      Height          =   420
      Left            =   6675
      TabIndex        =   12
      Top             =   2775
      Width           =   2550
   End
   Begin VB.CommandButton btnCancel 
      Caption         =   "&Cancel"
      Height          =   390
      Left            =   7275
      TabIndex        =   11
      Top             =   6075
      Width           =   1365
   End
   Begin VB.CommandButton btnOK 
      Caption         =   "&OK"
      Height          =   390
      Left            =   10125
      TabIndex        =   10
      Top             =   6075
      Width           =   1365
   End
   Begin VB.ComboBox apptype 
      Height          =   315
      ItemData        =   "ProjectManager.frx":2374E
      Left            =   2100
      List            =   "ProjectManager.frx":2375E
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   3300
      Width           =   3600
   End
   Begin VB.CheckBox Configed 
      Caption         =   "Generate Application Configuration"
      Height          =   195
      Left            =   2100
      TabIndex        =   7
      Top             =   1425
      Value           =   1  'Checked
      Width           =   3375
   End
   Begin VB.CommandButton langBrowse 
      Caption         =   "..."
      Height          =   285
      Left            =   5250
      TabIndex        =   6
      Top             =   2850
      Width           =   450
   End
   Begin VB.TextBox langPath 
      Height          =   285
      Left            =   2100
      Locked          =   -1  'True
      TabIndex        =   5
      Text            =   "[STARTUPPATH]\Languages"
      Top             =   2850
      Width           =   3075
   End
   Begin VB.CheckBox MintAPIRefrenced 
      Caption         =   "Auto Refrence To MintAPI"
      Height          =   195
      Left            =   2325
      TabIndex        =   3
      Top             =   1050
      Value           =   1  'Checked
      Width           =   2475
   End
   Begin VB.CheckBox MintAPIClient 
      Caption         =   "MintAPI Client Application"
      Height          =   195
      Left            =   2100
      TabIndex        =   2
      Top             =   750
      Value           =   1  'Checked
      Width           =   2745
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   2100
      TabIndex        =   0
      Text            =   "MintAPI_Project"
      Top             =   300
      Width           =   2625
   End
   Begin VB.Line Line2 
      Visible         =   0   'False
      X1              =   6600
      X2              =   6600
      Y1              =   0
      Y2              =   4950
   End
   Begin VB.Line Line1 
      Visible         =   0   'False
      X1              =   2025
      X2              =   2025
      Y1              =   150
      Y2              =   5100
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "Project Description:"
      Height          =   195
      Left            =   4650
      TabIndex        =   20
      Top             =   345
      Width           =   1995
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Current Language:"
      Height          =   195
      Left            =   75
      TabIndex        =   15
      Top             =   3885
      Width           =   1995
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Application Type:"
      Height          =   195
      Left            =   75
      TabIndex        =   9
      Top             =   3360
      Width           =   1995
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Language Directory:"
      Height          =   195
      Left            =   75
      TabIndex        =   4
      Top             =   2910
      Width           =   1995
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Project Name:"
      Height          =   195
      Left            =   225
      TabIndex        =   1
      Top             =   345
      Width           =   1845
   End
End
Attribute VB_Name = "ProjectManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public VBInstance As VBIDE.VBE
Public Connect As Connect

Private Sub btnApply_Click()
On Error Resume Next
    If VBInstance.ActiveVBProject Is Nothing Then
        Call MsgBox("No project selected.")
        Exit Sub
    End If
    VBInstance.ActiveVBProject.Name = txtName.Text
    VBInstance.ActiveVBProject.Description = txtDesc.Text
    If MintAPIRefrenced.Value = CheckBoxConstants.vbChecked And MintAPIClient.Value = CheckBoxConstants.vbChecked Then
        Call VBInstance.ActiveVBProject.References.AddFromFile(MintAPI.MintAPIDllPath)
    End If
    VBInstance.ActiveVBProject.Type = apptype.ListIndex
    PojectProperties(VBInstance.ActiveVBProject).ConfigurationPlace = IIf(regConfig.Value = CheckBoxConstants.vbChecked, ConfigurationPlace.cpRegistry, cpFile)
    PojectProperties(VBInstance.ActiveVBProject).LanguageDirectory = langPath.Text
'    Exit Sub
'err:
'    Call MsgBox(err.Description)
End Sub
Private Sub btnCancel_Click()
    Call Unload(Me)
End Sub
Private Sub btnOK_Click()
    Call btnApply_Click
    Call Unload(Me)
End Sub

Private Sub Form_Load()
On Error Resume Next
MsgBox VBInstance.ActiveVBProject.FileName
MsgBox C.getp.GetDirectory("[STARTUPPATH]")
    If VBInstance.ActiveVBProject Is Nothing Then
        Call MsgBox("No project selected.")
        Exit Sub
    End If
    txtName.Text = VBInstance.ActiveVBProject.Name
    txtDesc.Text = VBInstance.ActiveVBProject.Description
    apptype.ListIndex = VBInstance.ActiveVBProject.Type
    regConfig.Value = IIf(C.getp.ConfigurationPlace = cpRegistry, vbChecked, vbUnchecked)
    langPath.Text = C.getp.LanguageDirectory
End Sub

Private Sub langBrowse_Click()
On Error GoTo err
    Dim p As Directory
    Set p = Directory(C.getp.GetDirectory(langPath.Text))
    MsgBox C.getp.GetDirectory(langPath.Text)
    MsgBox p.AbsolutePath & "    " & p.Validate
    If Not p.Validate Then
        Set p = CurrentDirectory
    End If
    langPath.Text = Directory.ChooseDirectory(hWnd, "Project Language Files Directory...", sfCUSTOM, , p, True, False, False).AbsolutePath
    Exit Sub
err:
    Call MsgBox(err.Description)
End Sub

Private Sub Languaged_Click()
    langPath.Enabled = Languaged.Value = CheckBoxConstants.vbChecked
    langBrowse.Enabled = Languaged.Value = CheckBoxConstants.vbChecked
End Sub
Private Sub Configed_Click()
    regConfig.Enabled = Configed.Value = CheckBoxConstants.vbChecked
    fileConfig.Enabled = Configed.Value = CheckBoxConstants.vbChecked
End Sub

Private Sub MintAPIClient_Click()
    If MintAPIClient.Value = CheckBoxConstants.vbChecked Then
        MintAPIRefrenced.Enabled = True
    Else
        MintAPIRefrenced.Enabled = False
    End If
End Sub

Private Sub btnEditConfig_Click()
    Connect.ShowConfigurationEditor
End Sub
Private Sub btnEditLang_Click()
On Error GoTo err
Dim p As String
    If InStr(1, langPath.Text, "[STARTUPPATH]") > 0 Then
        p = langPath.Text
    Else
        p = "[STARTUPPATH]\languages"
    End If
    Dim d As Directory
    Dim ex As Boolean
    Set d = Directory(langPath.Text)
    ex = d.Exists
    If Not ex Then
        If Not VBInstance.ActiveVBProject.Saved Then
            On Error GoTo errE
            Set d = Directory.ChooseDirectory(hWnd, "Where to locate translation files...", , , CurrentDirectory)
        Else
            Set d = Directory(C.getp.GetDirectory(p))
        End If
        ex = d.Exists
    End If
    If Languaged.Value <> CheckBoxConstants.vbChecked Then
        If MsgBox("Do you want to generate language file for application?", vbYesNoCancel Or vbInformation) <> vbYes Then Exit Sub
        Languaged.Value = CheckBoxConstants.vbChecked
    End If
    If Not ex Then
        Call CreateLanguageFolder(d.AbsolutePath)
    End If
    langPath.Text = d.AbsolutePath
    Call C.ShowLanguageEditor(d.AbsolutePath)
    Exit Sub
err:
    MsgBox err.Description
errE:
End Sub
