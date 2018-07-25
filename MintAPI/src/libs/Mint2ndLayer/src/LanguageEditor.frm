VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmLanguageEditor 
   Caption         =   "Language Editor"
   ClientHeight    =   6975
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   9855
   BeginProperty Font 
      Name            =   "Consolas"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00404040&
   Icon            =   "LanguageEditor.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6975
   ScaleWidth      =   9855
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmButtons 
      Height          =   6495
      Left            =   9360
      TabIndex        =   1
      Top             =   0
      Width           =   495
      Begin VB.Image btnEdit 
         Height          =   240
         Left            =   120
         Picture         =   "LanguageEditor.frx":1082
         ToolTipText     =   "Edit Selected Language Record..."
         Top             =   1680
         Width           =   240
      End
      Begin VB.Image btnFindReplace 
         Height          =   240
         Left            =   120
         Picture         =   "LanguageEditor.frx":12B3
         ToolTipText     =   "Search Records..."
         Top             =   240
         Width           =   240
      End
      Begin VB.Image btnRemove 
         Height          =   240
         Left            =   120
         Picture         =   "LanguageEditor.frx":1501
         ToolTipText     =   "Remove Selected Language Record..."
         Top             =   1200
         Width           =   240
      End
      Begin VB.Image btnAdd 
         Height          =   240
         Left            =   120
         Picture         =   "LanguageEditor.frx":1722
         ToolTipText     =   "Add New Language Record..."
         Top             =   720
         Width           =   240
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grid 
      Height          =   6615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   11668
      _Version        =   393216
      Cols            =   3
      BackColor       =   16777215
      BackColorBkg    =   -2147483636
      SelectionMode   =   1
      AllowUserResizing=   3
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
         Caption         =   "Save &As..."
         Shortcut        =   ^D
      End
      Begin VB.Menu m_file_sep1 
         Caption         =   "-"
      End
      Begin VB.Menu m_file_import 
         Caption         =   "&Import..."
         Shortcut        =   ^I
      End
      Begin VB.Menu m_file_export 
         Caption         =   "&Export..."
         Shortcut        =   ^E
      End
      Begin VB.Menu m_file_sep2 
         Caption         =   "-"
      End
      Begin VB.Menu m_file_prop 
         Caption         =   "&Properties..."
         Shortcut        =   ^P
      End
      Begin VB.Menu m_file_sep3 
         Caption         =   "-"
      End
      Begin VB.Menu m_file_recent 
         Caption         =   "&Recent Language Files"
         Begin VB.Menu m_file_recent_items 
            Caption         =   "null"
            Index           =   0
         End
      End
      Begin VB.Menu m_file_sep4 
         Caption         =   "-"
      End
      Begin VB.Menu m_file_exit 
         Caption         =   "&Exit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu m_edit 
      Caption         =   "&Edit"
      Begin VB.Menu m_edit_add 
         Caption         =   "&Add New Record"
         Shortcut        =   ^V
      End
      Begin VB.Menu m_edit_removesel 
         Caption         =   "&Remove Selected Record"
         Shortcut        =   ^R
      End
      Begin VB.Menu m_edit_editsel 
         Caption         =   "&Edit Selected Record... (Double Click)"
         Shortcut        =   ^B
      End
      Begin VB.Menu m_edit_sep1 
         Caption         =   "-"
      End
      Begin VB.Menu m_edit_findreplace 
         Caption         =   "&Find And Replace..."
         Shortcut        =   ^F
      End
   End
   Begin VB.Menu m_settings 
      Caption         =   "&Settings"
      Begin VB.Menu m_set_preference 
         Caption         =   "&Preference..."
         Shortcut        =   ^K
      End
   End
   Begin VB.Menu m_about 
      Caption         =   "&About"
      Begin VB.Menu m_about_content_hlp 
         Caption         =   "&Content... (External HLP)"
         Shortcut        =   {F1}
      End
      Begin VB.Menu m_file_content 
         Caption         =   "&Content... (External HTML)"
         Shortcut        =   ^{F1}
      End
      Begin VB.Menu m_about_sep1 
         Caption         =   "-"
      End
      Begin VB.Menu m_about_about 
         Caption         =   "&About MintAPI Language Editor..."
         Shortcut        =   {F2}
      End
      Begin VB.Menu m_about_mintapi 
         Caption         =   "&About MintAPI..."
         Shortcut        =   {F4}
      End
   End
End
Attribute VB_Name = "frmLanguageEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const FILEFORMATS As String = "All Known Language Files (*.lng;*.lang;*.tr;*.translation;*.mlang;*.mintlng;*.mintlang;mintlanguage)|Language Files (*.lng;*.lang;*.tr;*.translation)|Mint Language Files (*.mlang;*.mintlng;*.mintlang;mintlanguage)|All Files (*)"
Const TITLE As String = "Language Editor"

Dim p As LanguageEditor

Dim lng As New Language

Dim p_edt_KeyBuffer As String
Dim p_edt_TrBuffer As String

Dim p_edt_lang_Name As String
Dim p_edt_lang_Key As String
Dim p_edt_lang_Description As String
Dim p_edt_lang_RightToLeft As Boolean
Dim p_edt_lang_Region As String

Dim p_edt_acceModify As Boolean
Dim iscalledbefore_boolean As Boolean


Dim lang_Name As String
Dim lang_Key As String
Dim lang_Description As String
Dim lang_RightToLeft As Boolean
Dim lang_Region As String

Dim last_savepath As String
Dim isSaved As Boolean

Private Sub InitializeGrid()
    Call ClearGrid
End Sub
Private Sub ReLoadGridFromLanguage()
    Call ClearGrid
    Dim i As Long, recCount As Long
    recCount = SpecialMethods.CountLanguageRecords(lng)
    grid.Rows = recCount + 1
    For i = 0 To recCount - 1
        grid.Row = i + 1
        grid.Col = 0
        grid.Text = (i + 1)
        grid.Col = 1
        grid.Text = SpecialMethods.LanguageRecordKey(lng, i)
        grid.Col = 2
        grid.Text = SpecialMethods.LanguageRecordTranslation(lng, i)
    Next
End Sub
Private Sub AppendRecordToGrid(Key As String, Translation As String)
    Dim rowsCount As Long
    If Not iscalledbefore_boolean Then
        rowsCount = 1
        iscalledbefore_boolean = True
    Else
        rowsCount = grid.Rows
        grid.Rows = rowsCount + 1
    End If
    grid.Row = rowsCount
    grid.Col = 0
    grid.Text = rowsCount
    grid.Col = 1
    grid.Text = Key
    grid.Col = 2
    grid.Text = Translation
End Sub
Private Sub RemoveRecordFromGrid(Index As Long)
    Dim i As Long, gridRowsCount As Long
    gridRowsCount = grid.Rows
    If Index < 1 Then
        Call MsgBox("Can't remove header row.", vbExclamation)
        Exit Sub
    End If
    If Index = 1 And gridRowsCount = 2 Then
        iscalledbefore_boolean = False
        grid.Row = 1
        grid.Col = 1
        grid.Text = ""
        grid.Col = 2
        grid.Text = ""
        Exit Sub
    End If
    Call grid.RemoveItem(Index)
    grid.Col = 0
    For i = Index To gridRowsCount - 2
        grid.Row = i
        grid.Text = i
    Next
End Sub
Public Sub SetEditorVariables(Modified As Boolean, Key As String, Translation As String)
    p_edt_acceModify = Modified
    p_edt_KeyBuffer = Key
    p_edt_TrBuffer = Translation
End Sub
Public Sub SetEditorProperties(Modified As Boolean, Name As String, ShortName As String, Region As String, Description As String, RightToLeft As Boolean)
    p_edt_acceModify = Modified
    p_edt_lang_Name = Name
    p_edt_lang_Key = ShortName
    p_edt_lang_Description = Description
    p_edt_lang_Region = Region
    p_edt_lang_RightToLeft = RightToLeft
End Sub
Public Sub FindReplace(ButtonType As ButtonType, FindWhat As String, ReplaceWith As String, SearchMode As SearchMode, MatchCase As Boolean)
    p_edt_acceModify = True
End Sub
Public Sub AppendRecord(Key As String, Translation As String)
    On Error GoTo Err
    Call SpecialMethods.AppendLanguageRecord(lng, Key, Translation)
    Call AppendRecordToGrid(Key, Translation)
    isSaved = False
    Call Edited
    Exit Sub
Err:
    Call MsgBox(Err.Description, vbInformation)
End Sub
Public Sub RemoveRecord(Index As Long)
On Error GoTo Err
    Dim Key As String
    grid.Row = Index
    grid.Col = 1
    Key = grid.Text
    If Key <> "" Then
        Call SpecialMethods.RemoveLanguageRecord(lng, Key)
        isSaved = False
        Call Edited
    End If
    Call RemoveRecordFromGrid(Index)
    Exit Sub
Err:
    Call MsgBox("An error occured while trying to remove language record ,Note that you can't remove header row.")
End Sub

Private Sub ClearGrid()
    Call grid.Clear
    grid.Rows = 2
    grid.ColWidth(0) = 38 * 15
    grid.ColWidth(1) = 150 * 15
    grid.ColWidth(2) = 500 * 15
    grid.TextMatrix(0, 0) = "Index"
    grid.TextMatrix(0, 1) = "Key"
    grid.TextMatrix(0, 2) = "Translation"
    grid.TextMatrix(1, 0) = "1"
End Sub
Private Sub ClearRecords()
    isSaved = True
    Call Edited(True)
    Call ClearGrid
    Call SpecialMethods.RemoveLanguageAllRecords(lng)
End Sub

Private Sub Form_Load()
    Me.Caption = TITLE
    lang_Name = "English"
    lang_Key = "en"
    lang_Region = "United States"
    lang_RightToLeft = False
    lang_Description = "English (United States) | Global English Language."
    isSaved = True
    Call Edited(True)
    Call InitializeGrid
    Dim fWidth As Long, fHeight As Long
    Dim fLeft As Long, fTop As Long
    On Error GoTo Err
    fWidth = Environment.GetMintAPIVariable("frmMintAPI_LanguageEditor_Width", Me.Width).toLong
    fHeight = Environment.GetMintAPIVariable("frmMintAPI_LanguageEditor_Height", Me.Height).toLong
    fLeft = Environment.GetMintAPIVariable("frmMintAPI_LanguageEditor_Left", Me.Left).toLong
    fTop = Environment.GetMintAPIVariable("frmMintAPI_LanguageEditor_Top", Me.Top).toLong
Err:
    If fWidth < Me.Width Then fWidth = Me.Width
    If fHeight < Me.Height Then fHeight = Me.Height
    
    If fLeft < 0 Then fLeft = 0
    If fTop < 0 Then fTop = 0
    
    If fLeft + fWidth > Screen.Width Then _
        fLeft = Screen.Width - fWidth

    If fTop + fHeight > Screen.Height Then _
        fTop = Screen.Height - fHeight
        
    Call Me.Move(fLeft, fTop, fWidth, fHeight)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Not p Is Nothing Then
        Dim Cancel1 As Boolean
        Cancel1 = False
        Call p.ClosingForm(Cancel1)
        Cancel = IIf(Cancel1, True, False)
    End If
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If frmButtons.Visible Then
        Call frmButtons.Move(ScaleWidth - frmButtons.Width, -90, frmButtons.Width, ScaleHeight + 75)
        Call grid.Move(0, 0, ScaleWidth - frmButtons.Width, ScaleHeight)
    Else
        Call grid.Move(0, 0, ScaleWidth, ScaleHeight)
    End If
End Sub

Public Sub SetLeftPanelVisible(Visible As Boolean)
    frmButtons.Visible = Visible
    Call Form_Resize
End Sub

Friend Sub SetParent(Parent As LanguageEditor)
    Set p = Parent
End Sub

Public Sub ShowFindReplace()
    Dim frmFind As New edt_find_replace
    Call frmFind.SetParent(Me)
    frmFind.Caption = "Find And Replace Languages..."
    frmFind.NoteInfo = "Find and replace language records in language editor."
    frmFind.txtFindWhat.Text = ""
    p_edt_acceModify = False
    Call ShowFormModally(frmFind, Me)
    If p_edt_acceModify Then _
        Call frmFind.SetFocus
    Set frmFind = Nothing
End Sub
Public Sub ShowProperties()
    Dim lngProps As New edt_Language_Settings
    Call lngProps.SetParent(Me)
    lngProps.Caption = "Language File Properties..."
    p_edt_acceModify = False
    lngProps.txtName = lang_Name
    lngProps.txtShortName = lang_Key
    lngProps.txtRegion = lang_Region
    lngProps.txtDescription = lang_Description
    lngProps.chkRightToLeft.Value = IIf(lang_RightToLeft, 1, 0)
    Call ShowFormModally(lngProps, Me)
    If p_edt_acceModify Then
        lang_Name = p_edt_lang_Name
        lang_Key = p_edt_lang_Key
        lang_Description = p_edt_lang_Description
        lang_Region = p_edt_lang_Region
        lang_RightToLeft = p_edt_lang_RightToLeft
    End If
    Set lngProps = Nothing
End Sub
Private Sub AddForm()
    Dim lngModifier As New edt_Language_Modify
    Call lngModifier.SetParent(Me)
    lngModifier.Caption = "Add Language Record..."
    lngModifier.NoteInfo = "Enter a key and then press OK button."
    p_edt_acceModify = False
    Call ShowFormModally(lngModifier, Me)
    If p_edt_acceModify Then
        Call AppendRecord(p_edt_KeyBuffer, p_edt_TrBuffer)
    End If
    Set lngModifier = Nothing
End Sub
Private Sub EditForm()
    If SpecialMethods.CountLanguageRecords(lng) <= 0 Then Exit Sub
    Dim recIndex As Long
    recIndex = grid.Row
    Dim lngModifier As New edt_Language_Modify
    Call lngModifier.SetParent(Me)
    Call lngModifier.SetEdit
    grid.Col = 2
    lngModifier.TranslationText = grid.Text
    grid.Col = 1
    lngModifier.KeyText = grid.Text
    If grid.Text <> "" Then _
        lngModifier.txtKey.Locked = True
    lngModifier.Caption = "Edit Language Record..."
    lngModifier.NoteInfo = "Enter a key and then press OK button."
    p_edt_acceModify = False
    Call ShowFormModally(lngModifier, Me)
    If p_edt_acceModify Then
        grid.Col = 2
        grid.Text = p_edt_TrBuffer
        grid.Col = 1
        If grid.Text = "" Then
            Call SpecialMethods.AppendLanguageRecord(lng, p_edt_KeyBuffer, p_edt_TrBuffer)
        Else
            Call SpecialMethods.ChangeLanguageRecord(lng, grid.Text, p_edt_TrBuffer)
        End If
    End If
    Set lngModifier = Nothing
End Sub
Private Sub RemoveForm()
    If MsgBox("Do you really want to remove selected record?", vbExclamation + vbYesNo, "Remove Warning...") <> vbYes Then Exit Sub
    Dim Index As Long
    Index = grid.RowSel
    'grid.
    If Index <= 0 Then Exit Sub
    Call RemoveRecord(Index)
End Sub

Private Sub frmButtons_Click()
    MsgBox SpecialMethods.CountLanguageRecords(lng)
End Sub

'####Remove All Records
Private Sub m_file_new_Click()
    If Not CheckSave Then Exit Sub
    Call ClearGrid
    Call ClearRecords
End Sub

Private Sub Edited(Optional State As Boolean = False)
    If last_savepath = "" Then
        Me.Caption = TITLE
        Exit Sub
    End If
    If State Then
        Me.Caption = TITLE & " - " & last_savepath
    Else
        Me.Caption = TITLE & " - " & last_savepath & " (*unsaved)"
    End If
End Sub

Public Sub SaveTo(Path As String)
On Error GoTo Err
    Dim lang As Language
    Set lang = SpecialMethods.CreateLanguage(lang_Name, lang_Key, lang_Region, lang_Description, lang_RightToLeft)
    Call SpecialMethods.CopyLanguageRecordsFromAnother(lang, lng)
    Call lang.CompileTranslationFile(Path)
    
    last_savepath = Path
    isSaved = True
    Call Edited(True)
    Exit Sub
Err:
    Call MsgBox("An error occured while trying to save language file." & vbCrLf & _
    "Original Error: " & Err.Description, vbCritical)
End Sub
Public Function Save() As Boolean
    If isSaved Then Exit Function
    Dim save_path As String
    If last_savepath = "" Then
        On Error GoTo Err
        save_path = File.ChooseFile(Me.hWnd, fdSave, "Save Translation File...", FILEFORMATS, "", CurrentDirectory).AbsolutePath
        Call Directory.SetCurrentDirectoryS(File(save_path).Location) 'or File(...).Directory.AbsolutePath
    Else
        save_path = last_savepath
    End If
    Call SaveTo(save_path)
    Save = True
    Exit Function
Err:
    Save = False
End Function
Public Function SaveAs() As Boolean
    On Error GoTo Err
    Dim save_aspath As String
    save_aspath = File.ChooseFile(Me.hWnd, fdSave, "Save Translation File...", FILEFORMATS, "", CurrentDirectory).AbsolutePath
    Call Directory.SetCurrentDirectoryS(File(save_aspath).Location) 'or File(...).Directory.AbsolutePath
    Call SaveTo(save_aspath)
    SaveAs = True
    Exit Function
Err:
    SaveAs = False
End Function
Public Function CheckSave() As Boolean
    If Not isSaved Then
        Dim msgRetVal As VbMsgBoxResult
        msgRetVal = MsgBox("Current language file not saved ,Do you want to save it?", vbExclamation + vbYesNoCancel, "Save File")
        If msgRetVal = vbYes Then _
            If Not Save Then CheckSave = CheckSave
        If msgRetVal = vbCancel Then CheckSave = False: Exit Function
    End If
    CheckSave = True
End Function

Private Sub m_file_open_Click()
    If Not CheckSave Then Exit Sub
    Dim open_path As String
    
    On Error GoTo operationCanceled
    open_path = File.ChooseFile(Me.hWnd, fdOpen, "Open Translation File...", FILEFORMATS, "", CurrentDirectory).AbsolutePath
    On Error GoTo Err
    Call Directory.SetCurrentDirectoryS(File(open_path).Location)
    
    Set lng = Language(open_path).Load
    lang_Name = lng.Name
    lang_Key = lng.ShortName
    lang_Region = lng.Region
    lang_RightToLeft = lng.RightToLeft
    lang_Description = lng.Description
    
    Call ReLoadGridFromLanguage
    
    last_savepath = open_path
    Exit Sub
Err:
    Call MsgBox("An error occured while trying to load language file." & vbCrLf & _
    "Original Error: " & Err.Description, vbCritical)
operationCanceled:
End Sub

Private Sub m_file_prop_Click()
    Call ShowProperties
End Sub

Private Sub m_file_save_Click()
    Call Save
End Sub

Private Sub m_file_saveas_Click()
    Call SaveAs
End Sub

Private Sub m_file_exit_Click()
    If Not CheckSave Then Exit Sub
    Call Unload(Me)
End Sub

Private Sub m_edit_findreplace_Click()
    Call ShowFindReplace
End Sub

Private Sub btnFindReplace_Click()
    Call ShowFindReplace
End Sub
Private Sub btnAdd_Click()
    Call AddForm
End Sub
Private Sub m_edit_add_Click()
    Call AddForm
End Sub
Private Sub btnEdit_Click()
    Call EditForm
End Sub
Private Sub m_edit_editsel_Click()
    Call EditForm
End Sub
Private Sub btnRemove_Click()
    Call RemoveForm
End Sub
Private Sub m_edit_removesel_Click()
    Call RemoveForm
End Sub

Private Sub grid_DblClick()
    Call EditForm
    If SpecialMethods.CountLanguageRecords(lng) <= 0 Then Call AddForm
End Sub

Private Sub m_about_mintapi_Click()
    Call SpecialMethods.AboutMintAPI(True)
End Sub

Private Sub m_about_about_Click()
    Call frmAbout.Show(1, Me)
End Sub
