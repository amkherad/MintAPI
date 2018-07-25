Attribute VB_Name = "envMethods"
'@PROJECT_LICENSE

Option Explicit
Option Base 0
Const CLASSID As String = "envMethods"

Private Const API_envMethods_NIF_ICON = &H2
Private Const API_envMethods_NIF_MESSAGE = &H1
Private Const API_envMethods_NIF_TIP = &H4
Private Const API_envMethods_NIM_ADD = &H0
Private Const API_envMethods_NIM_MODIFY = &H1
Private Const API_envMethods_NIM_DELETE = &H2
Private Const NOTIFYICONDATATOOLTIPSIZE As Long = 64

Public Const API_OFN_ALLOWMULTISELECT = &H200
Public Const API_OFN_CREATEPROMPT = &H2000
Public Const API_OFN_ENABLEHOOK = &H20
Public Const API_OFN_ENABLETEMPLATE = &H40
Public Const API_OFN_ENABLETEMPLATEHANDLE = &H80
Public Const API_OFN_ENABLESIZING = &H800000
Public Const API_OFN_EXPLORER = &H80000
Public Const API_OFN_FILEMUSTEXIST = &H1000
Public Const API_OFN_EXTENSIONDIFFERENT = &H400
Public Const API_OFN_HIDEREADONLY = &H4
Public Const API_OFN_LONGNAMES = &H200000
Public Const API_OFN_NOCHANGEDIR = &H8
Public Const API_OFN_NODEREFERENCELINKS = &H100000
Public Const API_OFN_NOLONGNAMES = &H40000
Public Const API_OFN_NOREADONLYRETURN = &H8000&
Public Const API_OFN_NONETWORKBUTTON = &H20000
Public Const API_OFN_NOTESTFILECREATE = &H10000
Public Const API_OFN_NOVALIDATE = &H100
Public Const API_OFN_OVERWRITEPROMPT = &H2
Public Const API_OFN_PATHMUSTEXIST = &H800
Public Const API_OFN_READONLY = &H1
Public Const API_OFN_SHAREAWARE = &H4000
Public Const API_OFN_SHAREFALLTHROUGH = 2
Public Const API_OFN_SHARENOWARN = 1
Public Const API_OFN_SHAREWARN = 0
Public Const API_OFN_SHOWHELP = &H10

Private Const BIF_RETURNONLYFSDIRS   As Long = &H1          ' For finding a folder to start document searching
Private Const BIF_DONTGOBELOWDOMAIN  As Long = &H2          ' For starting the Find Computer
Private Const BIF_STATUSTEXT         As Long = &H4          ' Top of the dialog has 2 lines of text for BROWSEINFO.lpszTitle and one line if
                                                           ' this flag is set.  Passing the message BFFM_SETSTATUSTEXTA to the hwnd can set the
                                                           ' rest of the text.  This is not used with BIF_USENEWUI and BROWSEINFO.lpszTitle gets
                                                           ' all three lines of text.
Private Const BIF_RETURNFSANCESTORS  As Long = &H8
Private Const BIF_EDITBOX            As Long = &H10         ' Add an editbox to the dialog
Private Const BIF_VALIDATE           As Long = &H20         ' insist on valid result (or CANCEL)
Private Const BIF_NEWDIALOGSTYLE     As Long = &H40         ' Use the new dialog layout with the ability to resize
                                                           ' Caller needs to call OleInitialize() before using this API
Private Const BIF_USENEWUI           As Long = (BIF_NEWDIALOGSTYLE Or BIF_EDITBOX)
Private Const BIF_BROWSEINCLUDEURLS  As Long = &H80         ' Allow URLs to be displayed or entered. (Requires BIF_USENEWUI)
Private Const BIF_UAHINT             As Long = &H100        ' Add a UA hint to the dialog, in place of the edit box. May not be combined with BIF_EDITBOX
Private Const BIF_NONEWFOLDERBUTTON  As Long = &H200        ' Do not add the "New Folder" button to the dialog.  Only applicable with BIF_NEWDIALOGSTYLE.
Private Const BIF_NOTRANSLATETARGETS As Long = &H400        ' don't traverse target as shortcut

Private Const BIF_BROWSEFORCOMPUTER  As Long = &H1000       ' Browsing for Computers.
Private Const BIF_BROWSEFORPRINTER   As Long = &H2000       ' Browsing for Printers
Private Const BIF_BROWSEINCLUDEFILES As Long = &H4000       ' Browsing for Everything
Private Const BIF_SHAREABLE          As Long = &H8000       ' sharable resources displayed (remote shares, requires BIF_USENEWUI)
Private Const BIF_BROWSEFILEJUNCTIONS As Long = &H10000     ' allow folder junctions like zip files and libraries to be browsed

Public Enum API_DialogType 'must be Public
    OpenDialog = 0
    SaveDialog = 1
End Enum
Private Type API_envMethods_NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * NOTIFYICONDATATOOLTIPSIZE
End Type
Private Type API_OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    Flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type
Private Type API_BROWSEINFO
    hOwner As Long
    pidlRoot As Long
    pszDisplayName As String
    lpszTitle As String
    ulFlags As Long
    lpfn As Long
    lParam As Long
    iImage As Long
End Type

Private Declare Function API_envMethods_LoadCursorFromFile Lib "user32" Alias "LoadCursorFromFileA" (ByVal lpFileName As String) As Long
Private Declare Function API_envMethods_Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As API_envMethods_NOTIFYICONDATA) As Long
Private Declare Function API_envMethods_LoadIcon Lib "user32" Alias "LoadIconA" (ByVal hInstance As Long, ByVal lpIconName As String) As Long
Private Declare Function API_envMethods_LoadIconFromFile Lib "user32" Alias "LoadIconFromFileA" (ByVal lpIconName As String) As Long
Private Declare Function API_envMethods_GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Function API_envMethods_SHBrowseForFolder Lib "shell32" (ByRef lpbi As API_BROWSEINFO) As Long
'Private Declare Function API_envMethods_SHGetPathFromIDList Lib "shell32" (ByRef pidl As API_CITEMIDLIST, ByVal pszPath As String) As Long

Private Declare Function API_envMethods_GetOpenFileName Lib "comdlg32" Alias "GetOpenFileNameA" (pOpenfilename As API_OPENFILENAME) As Long
Private Declare Function API_envMethods_GetSaveFileName Lib "comdlg32" Alias "GetSaveFileNameA" (pOpenfilename As API_OPENFILENAME) As Long


Dim inited As Boolean
Dim hInstance As Long

Public Sub Initialize()
    If inited Then Exit Sub
    'Call STDCONSTS.Initialize
    Call Exceptions.Initialize
    inited = True
End Sub
Public Sub Dispose(Optional ByVal Force As Boolean = False)
    If Not inited Then Exit Sub
    Call STDCONSTS.Dispose(Force)
    inited = False
End Sub

Public Function LoadIcon(Path As String) As Long
    If Dir(Path) = "" Then throw FileNotFoundException
    LoadIcon = API_envMethods_LoadCursorFromFile(Path)
    If LoadIcon = 0 Then throw SystemCallFailureException
End Function
Public Function CreateTrayIcon(hwnd As Long, hIcon As Long, ToolTip As String, uID As Long, CallbackMessage As Long)
    If Len(ToolTip) > NOTIFYICONDATATOOLTIPSIZE Then throw StringLengthOverloadException
    Dim tr As API_envMethods_NOTIFYICONDATA
    tr.hIcon = hIcon
    tr.hwnd = hwnd
    tr.szTip = ToolTip & Chr(0)
    tr.uCallbackMessage = CallbackMessage
    tr.uID = uID
    Dim FL As Long
    FL = IIf(hIcon = 0, 0, API_envMethods_NIF_ICON)
    FL = FL Or IIf(ToolTip = "", 0, API_envMethods_NIF_TIP)
    FL = FL Or IIf(CallbackMessage = 0, 0, API_envMethods_NIF_MESSAGE)
    tr.uFlags = FL
    tr.cbSize = Len(tr)
    CreateTrayIcon = API_envMethods_Shell_NotifyIcon(API_envMethods_NIM_ADD, tr)
End Function
Public Sub DestroyTrayIcon(hwnd As Long)
    Dim tr As API_envMethods_NOTIFYICONDATA
    tr.hwnd = hwnd
    tr.uFlags = 0
    tr.cbSize = Len(tr)
    Call API_envMethods_Shell_NotifyIcon(API_envMethods_NIM_DELETE, tr)
End Sub

Public Function Dialogs_BrowseFile(DialogType As API_DialogType, hWndParent As Long, Title As String, InitialDirectory As String, Filter As String, Flags As Long) As String
    Dim f As API_OPENFILENAME
    f.lStructSize = Len(f)
    f.hwndOwner = hWndParent
    f.lpstrFilter = Filter
    'f.lpstrCustomFilter = ""
    f.nMaxCustFilter = 0
    f.nFilterIndex = -1
    f.lpstrFile = String(LARGELPSTR, Chr(0))
    f.nMaxFile = LARGELPSTR
    f.hInstance = API_Instance
    f.lpstrInitialDir = InitialDirectory
    f.lpstrTitle = Title
    f.Flags = API_OFN_EXPLORER Or API_OFN_ENABLESIZING Or API_OFN_HIDEREADONLY Or Flags
    If DialogType = OpenDialog Then
        '//fileName.Flags = fileName.Flags | OFN_FILEMUSTEXIST + OFN_PATHMUSTEXIST;
        If API_envMethods_GetOpenFileName(f) = 0 Then throw OperationCanceledException
    Else
        '// dialogType == SAVE_FILE_DIALOG)
        If API_envMethods_GetSaveFileName(f) = 0 Then throw OperationCanceledException
    End If
    Dialogs_BrowseFile = GetLPSTR(f.lpstrFile) & Chr(0)
End Function
Public Function Dialogs_Browse(hWndParent As Long, Title As String, CreateNewButton As Boolean, Flags As Long)
'    Dim f As API_BROWSEINFO
'    f.hOwner = hWndParent
'    f.lpszTitle = Title
'    f.ulFlags = BIF_RETURNONLYFSDIRS Or BIF_STATUSTEXT Or Flags
'    If Not CreateNewButton Then
'        f.ulFlags = bi.ulFlags Or BIF_NONEWFOLDERBUTTON
'    End If
'
'    pidl = API_envMethods_SHBrowseForFolder(bi)
'
'    If API_envMethods_SHBrowseForFolder(f) = 0 Then throw OperationCanceledException
End Function
