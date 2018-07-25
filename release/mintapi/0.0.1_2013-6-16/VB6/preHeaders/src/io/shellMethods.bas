Attribute VB_Name = "shellMethods"
'@PROJECT_LICENSE

Option Explicit
Option Base 0
Const CLASSID As String = "shellMethods"

Private Const API_shellMethods_NIF_ICON = &H2
Private Const API_shellMethods_NIF_MESSAGE = &H1
Private Const API_shellMethods_NIF_TIP = &H4
Private Const API_shellMethods_NIM_ADD = &H0
Private Const API_shellMethods_NIM_MODIFY = &H1
Private Const API_shellMethods_NIM_DELETE = &H2
Private Const NOTIFYICONDATATOOLTIPSIZE As Long = 64

Public Const API_OFN_SHAREWARN = 0
Public Const API_OFN_SHARENOWARN = 1
Public Const API_OFN_READONLY = &H1
Public Const API_OFN_SHAREFALLTHROUGH = 2
Public Const API_OFN_OVERWRITEPROMPT = &H2
Public Const API_OFN_HIDEREADONLY = &H4
Public Const API_OFN_NOCHANGEDIR = &H8
Public Const API_OFN_SHOWHELP = &H10
Public Const API_OFN_ENABLEHOOK = &H20
Public Const API_OFN_ENABLETEMPLATE = &H40
Public Const API_OFN_ENABLETEMPLATEHANDLE = &H80
Public Const API_OFN_NOVALIDATE = &H100
Public Const API_OFN_ALLOWMULTISELECT = &H200
Public Const API_OFN_EXTENSIONDIFFERENT = &H400
Public Const API_OFN_PATHMUSTEXIST = &H800
Public Const API_OFN_FILEMUSTEXIST = &H1000
Public Const API_OFN_CREATEPROMPT = &H2000
Public Const API_OFN_SHAREAWARE = &H4000
Public Const API_OFN_NOREADONLYRETURN = &H8000&
Public Const API_OFN_NOTESTFILECREATE = &H10000
Public Const API_OFN_NONETWORKBUTTON = &H20000
Public Const API_OFN_NOLONGNAMES = &H40000
Public Const API_OFN_EXPLORER = &H80000
Public Const API_OFN_NODEREFERENCELINKS = &H100000
Public Const API_OFN_LONGNAMES = &H200000
Public Const API_OFN_ENABLESIZING = &H800000

Private Const WM_USER = &H400
Private Const BFFM_INITIALIZED = 1
Private Const BFFM_SELCHANGED = 2
Private Const BFFM_SETSTATUSTEXT = (WM_USER + 100)
Private Const BFFM_SETSELECTION = (WM_USER + 102)
Private Const BFFM_ENABLEOK = (WM_USER + 101)
Private Const BFFM_SETOKTEXT = (WM_USER + 105)

Public Const BIF_RETURNONLYFSDIRS   As Long = &H1          ' For finding a folder to start document searching
Public Const BIF_DONTGOBELOWDOMAIN  As Long = &H2          ' For starting the Find Computer
Public Const BIF_STATUSTEXT         As Long = &H4          ' Top of the dialog has 2 lines of text for BROWSEINFO.lpszTitle and one line if
                                                           ' this flag is set.  Passing the message BFFM_SETSTATUSTEXTA to the hwnd can set the
                                                           ' rest of the text.  This is not used with BIF_USENEWUI and BROWSEINFO.lpszTitle gets
                                                           ' all three lines of text.
Public Const BIF_RETURNFSANCESTORS  As Long = &H8
Public Const BIF_EDITBOX            As Long = &H10         ' Add an editbox to the dialog
Public Const BIF_VALIDATE           As Long = &H20         ' insist on valid result (or CANCEL)
Public Const BIF_NEWDIALOGSTYLE     As Long = &H40         ' Use the new dialog layout with the ability to resize
                                                           ' Caller needs to call OleInitialize() before using this API
Public Const BIF_USENEWUI           As Long = (BIF_NEWDIALOGSTYLE Or BIF_EDITBOX)
Public Const BIF_BROWSEINCLUDEURLS  As Long = &H80         ' Allow URLs to be displayed or entered. (Requires BIF_USENEWUI)
Public Const BIF_UAHINT             As Long = &H100        ' Add a UA hint to the dialog, in place of the edit box. May not be combined with BIF_EDITBOX
Public Const BIF_NONEWFOLDERBUTTON  As Long = &H200        ' Do not add the "New Folder" button to the dialog.  Only applicable with BIF_NEWDIALOGSTYLE.
Public Const BIF_NOTRANSLATETARGETS As Long = &H400        ' don't traverse target as shortcut

Public Const BIF_BROWSEFORCOMPUTER  As Long = &H1000       ' Browsing for Computers.
Public Const BIF_BROWSEFORPRINTER   As Long = &H2000       ' Browsing for Printers
Public Const BIF_BROWSEINCLUDEFILES As Long = &H4000       ' Browsing for Everything
Public Const BIF_SHAREABLE          As Long = &H8000       ' sharable resources displayed (remote shares, requires BIF_USENEWUI)
Public Const BIF_BROWSEFILEJUNCTIONS As Long = &H10000     ' allow folder junctions like zip files and libraries to be browsed

'BIF_RETURNONLYFSDIRS or BIF_DONTGOBELOWDOMAIN or BIF_STATUSTEXT or BIF_RETURNFSANCESTORS or BIF_EDITBOX
    'BIF_VALIDATE or BIF_NEWDIALOGSTYLE or BIF_USENEWUI or BIF_BROWSEINCLUDEURLS or BIF_UAHINT or BIF_NONEWFOLDERBUTTON
    'BIF_NOTRANSLATETARGETS or BIF_BROWSEFORCOMPUTER or BIF_BROWSEFORPRINTER or BIF_BROWSEINCLUDEFILES or BIF_SHAREABLE
    'BIF_BROWSEFILEJUNCTIONS
Public Enum API_DialogType 'must be Public
    OpenDialog = 0
    SaveDialog = 1
    FileSelectDialog = 2
End Enum
Public Type API_FileDialogReturn
    Path As String
    ReadOnlyCheckState As Boolean
    subFiles() As String
    subFilesCount As Long
End Type
Public Enum API_ShellDialogType
    ShellDialog_Sizable = 1
    ShellDialog_NewButton = 2
    ShellDialog_ShowFiles = 4
    ShellDialog_BrowseFileJunctions = 8
    ShellDialog_EditBox = &H10
    ShellDialog_Default = ShellDialog_Sizable Or ShellDialog_NewButton
End Enum
Private Type API_shellMethods_NOTIFYICONDATA
    cbSize As Long
    hWnd As Long
    uID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * NOTIFYICONDATATOOLTIPSIZE
End Type
Public Enum API_SpecialFolders
    sf_Custom = -1
    sf_Desktop = &H0
    sf_All = sf_Desktop
    sf_Programs = &H2
    sf_User_Start_Menu_Programs = sf_Programs
    sf_Controls = &H3
    sf_Printers = &H4
    sf_Personal = &H5
    sf_User_MyDocuments = sf_Personal
    sf_Favorites = &H6
    sf_Startup = &H7
    sf_User_Start_Menu_Programs_Startup = sf_Startup
    sf_Recent = &H8
    sf_SendTo = &H9
    sf_BitBucket = &HA
    sf_StartMenu = &HB
    sf_MyMusic = &HD
    sf_MyVideos = &HE
    sf_MyPictures = &H27
    sf_User_StartMenu = sf_StartMenu
    sf_User_Desktop = &H10
    sf_DesktopDirectory = sf_User_Desktop
    sf_Drives = &H11
    sf_MyComputer = sf_Drives
    sf_Network = &H12
    sf_All_Network = sf_Network
    sf_Nethood = &H13
    sf_Fonts = &H14
    sf_Templates = &H15
    sf_Common_StartMenu = &H16
    sf_Common_StartMenu_Programs = &H17
    sf_Common_StartMenu_Programs_Startup = &H18
    sf_Common_Desktop = &H19

    sf_ApplicationData = &H1A
    sf_PrintHood = &H1B
    sf_LocalApplicationData = &H1C
    sf_Common_Favorites = &H1F
    sf_Temp_InternetFiles = &H20
    sf_Cookies = &H21
    sf_History = &H22
    sf_Common_ApplicationData = &H23

    sf_Windows = &H24
    sf_System = &H25
    sf_Program_Files = &H26
    sf_User = &H28
    sf_Common_Templates = &H2D
    sf_ProgramFiles_CommonFiles = &H2B
    sf_Common_Documents = &H2E
    sf_Common_AdministrativeTools = &H2F
    sf_AdministrativeTools = &H30
    sf_Common_MyMusic = &H35
    sf_Common_MyPictures = &H36
    sf_Common_MyVideos = &H37
    sf_Resources = &H38
    sf_CDBurning = &H3B

    sf_Workgroup = &H3D
    sf_Network_Computers = sf_Workgroup
End Enum
Private Type API_OPENFILENAME
    lStructSize As Long
    hWndOwner As Long
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

Private Type SH_ITEM_ID
    cb As Long
    abID As Byte
End Type
Private Type ITEMIDLIST
    mkid As SH_ITEM_ID
End Type

Private Type OSVersionInfo
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128 ' Maintenance string For PSS usage
End Type

Private Declare Function API_shellMethods_SHSimpleIDListFromPath Lib "shell32" Alias "#162" (ByVal szPath As String) As Long
Private Declare Function API_shellMethods_SHGetPathFromIDList Lib "shell32" Alias "SHGetPathFromIDList" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Private Declare Function API_shellMethods_SHBrowseForFolder Lib "shell32" Alias "SHBrowseForFolder" (lpbi As API_BROWSEINFO) As Long
Private Declare Sub API_shellMethods_CoTaskMemFree Lib "ole32" Alias "CoTaskMemFree" (ByVal hMem As Long)
Private Declare Function API_shellMethods_SHGetSpecialFolderLocation Lib "shell32" Alias "SHGetSpecialFolderLocation" (ByVal hWndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long
Private Declare Sub API_shellMethods_OleInitialize Lib "ole32" Alias "OleInitialize" (pvReserved As Any)
Private Declare Function API_shellMethods_PathIsDirectory Lib "shlwapi" Alias "PathIsDirectoryA" (ByVal pszPath As String) As Long
Private Declare Function API_shellMethods_SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Private Declare Function API_shellMethods_SendMessage2 Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function API_shellMethods_GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVersionInfo) As Long

Private Declare Function API_shellMethods_GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Function API_shellMethods_LoadCursorFromFile Lib "user32" Alias "LoadCursorFromFileA" (ByVal lpFileName As String) As Long
Private Declare Function API_shellMethods_Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As API_shellMethods_NOTIFYICONDATA) As Long
Private Declare Function API_shellMethods_LoadIcon Lib "user32" Alias "LoadIconA" (ByVal hInstance As Long, ByVal lpIconName As String) As Long
Private Declare Function API_shellMethods_LoadIconFromFile Lib "user32" Alias "LoadIconFromFileA" (ByVal lpIconName As String) As Long

Private Declare Function API_shellMethods_GetOpenFileName Lib "comdlg32" Alias "GetOpenFileNameA" (pOpenfilename As API_OPENFILENAME) As Long
Private Declare Function API_shellMethods_GetSaveFileName Lib "comdlg32" Alias "GetSaveFileNameA" (pOpenfilename As API_OPENFILENAME) As Long

Private Declare Function API_shellMethods_ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long


Dim OK_BUTTON_TEXT As String 'Dialog_Browse
Dim m_CurrentDirectory As String 'Dialog_Browse

Dim inited As Boolean
Dim hInstance As Long

Public Sub Initialize()
    If inited Then Exit Sub
    Call baseConstants.Initialize
    Call baseExceptions.Initialize
    inited = True
End Sub
Public Sub Dispose(Optional ByVal Force As Boolean = False)
    If Not inited Then Exit Sub
    inited = False
End Sub

Public Function LoadIcon(Path As String) As Long
    If Dir(Path) = "" Then throw FileNotFoundException
    LoadIcon = API_shellMethods_LoadCursorFromFile(Path)
    If LoadIcon = 0 Then throw SystemCallFailureException
End Function
Public Function CreateTrayIcon(hWnd As Long, hIcon As Long, ToolTip As String, uID As Long, CallbackMessage As Long) As Long
    If Len(ToolTip) > NOTIFYICONDATATOOLTIPSIZE Then throw StringLengthOverloadException
    Dim tr As API_shellMethods_NOTIFYICONDATA
    tr.hIcon = hIcon
    tr.hWnd = hWnd
    tr.szTip = ToolTip & Chr(0)
    tr.uCallbackMessage = CallbackMessage
    tr.uID = uID
    Dim fl As Long
    fl = IIf(hIcon = 0, 0, API_shellMethods_NIF_ICON)
    fl = fl Or IIf(ToolTip = "", 0, API_shellMethods_NIF_TIP)
    fl = fl Or IIf(CallbackMessage = 0, 0, API_shellMethods_NIF_MESSAGE)
    tr.uFlags = fl
    tr.cbSize = Len(tr)
    CreateTrayIcon = API_shellMethods_Shell_NotifyIcon(API_shellMethods_NIM_ADD, tr)
End Function
Public Sub DestroyTrayIcon(tray_hWnd As Long)
    Dim tr As API_shellMethods_NOTIFYICONDATA
    tr.hWnd = tray_hWnd
    tr.uFlags = 0
    tr.cbSize = Len(tr)
    Call API_shellMethods_Shell_NotifyIcon(API_shellMethods_NIM_DELETE, tr)
End Sub


Public Function GetSpecialfolder(SpecialFolders As API_SpecialFolders) As String
    Dim R As Long
    Dim IDL As ITEMIDLIST
    'Get the special folder
    R = API_shellMethods_SHGetSpecialFolderLocation(100, SpecialFolders, IDL)
    If R = 0 Then
        'Create a buffer
        Dim Path As String
        Path = String(LARGELPSTR, Chr(0))
        'Get the path from the IDList
        R = API_shellMethods_SHGetPathFromIDList(ByVal IDL.mkid.cb, ByVal Path$)
        'Remove the unnecessary chr$(0)'s
        GetSpecialfolder = Left$(Path, InStr(Path, Chr$(0)) - 1)
        Exit Function
    End If
    throw Exception("Unknown SpecialFolder Value.")
End Function

Private Function SplitFileFilterbyMethod01(Filter As String) As String
    'Filter Format: AllFiles (*.*;*)   ->   AllFiles (*.*;*)|*.*;*
    If Filter = "" Then Exit Function
    Dim retVal As String
    
End Function

Public Function Dialogs_BrowseFile( _
        DialogType As API_DialogType, hWndParent As Long, Title As String, _
        InitialDirectory As String, Filter As String, FileName As String, _
        Flags As Long, Optional CheckPathExists As Boolean = True, _
        Optional CheckFilesExists As Boolean = True, _
        Optional AllowMultiSelect As Boolean = False, _
        Optional OpenAsReadOnly As Boolean = False, _
        Optional OverWritePrompt As Boolean = True, _
        Optional AllowChangingDirectory As Boolean = True) As API_FileDialogReturn

    Dim f As API_OPENFILENAME
    f.hWndOwner = hWndParent

    If Filter <> "" Then _
        f.lpstrFilter = SplitFileFilterbyMethod01(Filter)
    
    f.nMaxCustFilter = -1
    f.nFilterIndex = 0
    
    'f.lpstrFile = strFilter
    If AllowMultiSelect Then
        f.lpstrFile = String(MULTILPSTR, Chr(0))
        f.nMaxFile = MULTILPSTR
        f.lpstrFileTitle = String(LARGELPSTR, Chr(0))
        f.nMaxFileTitle = LARGELPSTR
    Else
        f.lpstrFile = String(LARGELPSTR, Chr(0))
        f.nMaxFile = LARGELPSTR
        f.lpstrFileTitle = String(LARGELPSTR, Chr(0))
        f.nMaxFileTitle = LARGELPSTR
    End If
    
    f.hInstance = App.hInstance
    f.lpstrInitialDir = InitialDirectory
    f.lpstrTitle = Title
    
    Dim gFlags As Long
    'gFlags = Flags Or API_OFN_NOCHANGEDIR
    If CheckPathExists Then gFlags = gFlags Or API_OFN_PATHMUSTEXIST
    If CheckFilesExists Then gFlags = gFlags Or API_OFN_FILEMUSTEXIST
    If AllowMultiSelect Then gFlags = gFlags Or API_OFN_ALLOWMULTISELECT
    If Not OpenAsReadOnly Then gFlags = gFlags Or API_OFN_HIDEREADONLY
    If OverWritePrompt Then gFlags = gFlags Or API_OFN_OVERWRITEPROMPT
    If Not AllowChangingDirectory Then gFlags = gFlags Or API_OFN_NOCHANGEDIR
    f.Flags = gFlags Or API_OFN_ENABLESIZING Or API_OFN_EXPLORER
    
    f.lStructSize = Len(f)
    If DialogType = OpenDialog Then
        If API_shellMethods_GetOpenFileName(f) = 0 Then throw OperationCanceledException
    Else
        If API_shellMethods_GetSaveFileName(f) = 0 Then throw OperationCanceledException
    End If
    
    If AllowMultiSelect Then
        On Error GoTo err1
        Dim string_n_null As String, fString As String
        Dim cIndex As Long, cNullStr As Long
        fString = f.lpstrFile
        cIndex = 1
        cNullStr = InStr(1, fString, Chr(0))
        If cNullStr > 0 Then
            string_n_null = Left(fString, cNullStr - 1)
            Dialogs_BrowseFile.Path = string_n_null
            Do
                cIndex = cNullStr + 1
                cNullStr = InStr(cIndex, fString, Chr(0))
                string_n_null = mID(fString, cIndex, (cNullStr - cIndex))
                If string_n_null = "" Then
                    Exit Do
                Else
                    ReDim Preserve Dialogs_BrowseFile.subFiles(Dialogs_BrowseFile.subFilesCount)
                    Dialogs_BrowseFile.subFiles(Dialogs_BrowseFile.subFilesCount) = string_n_null
                    Dialogs_BrowseFile.subFilesCount = Dialogs_BrowseFile.subFilesCount + 1
                End If
            Loop
        End If
err1:
    Else
        Dialogs_BrowseFile.Path = mID(f.lpstrFile, 1, InStr(1, f.lpstrFile, Chr(0)) - 1)
    End If
End Function
'BIF_RETURNONLYFSDIRS or BIF_DONTGOBELOWDOMAIN or BIF_STATUSTEXT or BIF_RETURNFSANCESTORS or BIF_EDITBOX
'BIF_VALIDATE or BIF_NEWDIALOGSTYLE or BIF_USENEWUI or BIF_BROWSEINCLUDEURLS or BIF_UAHINT or BIF_NONEWFOLDERBUTTON
'BIF_NOTRANSLATETARGETS or BIF_BROWSEFORCOMPUTER or BIF_BROWSEFORPRINTER or BIF_BROWSEINCLUDEFILES or BIF_SHAREABLE
'BIF_BROWSEFILEJUNCTIONS

Private Function isNT2000XP() As Boolean
    Dim lpv As OSVersionInfo
    lpv.dwOSVersionInfoSize = Len(lpv)
    Call API_shellMethods_GetVersionEx(lpv)
    If lpv.dwPlatformId = 2 Then
        isNT2000XP = True
    Else
        isNT2000XP = False
    End If
End Function
Private Function isME2KXP() As Boolean
    Dim lpv As OSVersionInfo
    lpv.dwOSVersionInfoSize = Len(lpv)
    Call API_shellMethods_GetVersionEx(lpv)
    If ((lpv.dwPlatformId = 2) And (lpv.dwMajorVersion >= 5)) Or _
    ((lpv.dwPlatformId = 1) And (lpv.dwMajorVersion >= 4) And (lpv.dwMinorVersion >= 90)) Then
    isME2KXP = True
Else
    isME2KXP = False
End If
End Function
Private Function GetPIDLFromPath(sPath As String) As Long
    ' Return the pidl to the path supplied b
    '     y calling the undocumented API #162
    If isNT2000XP Then
        GetPIDLFromPath = API_shellMethods_SHSimpleIDListFromPath(StrConv(sPath, vbUnicode))
    Else
        GetPIDLFromPath = API_shellMethods_SHSimpleIDListFromPath(sPath)
    End If
End Function

Public Function Dialogs_Browse(Optional ByVal OwnerForm As Long = 0, Optional ByVal Title As String = "", _
            Optional ByVal RootDir As API_SpecialFolders = sf_All, Optional ByVal CustomRootDir As String = "", _
            Optional ByVal StartDir As String = "", Optional ByVal NewStyle As Boolean = True, _
            Optional ByVal IncludeFiles As Boolean = False, Optional ByVal ShowNewFolderButton As Boolean = True, _
            Optional ByVal OkButtonText As String = "", Optional ByVal Flags As Long = 0) As String
    Dim lpIDList As Long, sBuffer As String, tBrowseInfo As API_BROWSEINFO, clRoot As Boolean


    If Len(OkButtonText) > 0 Then
        OK_BUTTON_TEXT = OkButtonText
    Else
        OK_BUTTON_TEXT = vbNullString
    End If
    clRoot = False

    If RootDir = sf_Custom Then


        If Len(CustomRootDir) > 0 Then


            If (API_shellMethods_PathIsDirectory(CustomRootDir) And (Left$(CustomRootDir, 2) <> "\\")) Or (Left$(CustomRootDir, 2) = "\\") Then
                tBrowseInfo.pidlRoot = GetPIDLFromPath(CustomRootDir)
                'SHILCreateFromPath StrPtr(CustomRootDir
                '     ), tBrowseInfo.pidlRoot, ByVal 0&
                clRoot = True
            Else
                tBrowseInfo.pidlRoot = GetSpecialFolderID(sf_MyComputer)
            End If
        Else
            tBrowseInfo.pidlRoot = GetSpecialFolderID(sf_All)
        End If
    Else
        tBrowseInfo.pidlRoot = GetSpecialFolderID(RootDir)
    End If

    If (Len(StartDir) > 0) Then
        m_CurrentDirectory = StartDir & vbNullChar
    Else
        m_CurrentDirectory = vbNullChar
    End If

    If Len(Title) > 0 Then
        tBrowseInfo.lpszTitle = Title
    Else
        tBrowseInfo.lpszTitle = "Select A Directory"
    End If
    tBrowseInfo.lpfn = GetAddressOfFunction(AddressOf Browse_CallbackProc)
    tBrowseInfo.ulFlags = BIF_RETURNONLYFSDIRS
    If IncludeFiles Then tBrowseInfo.ulFlags = tBrowseInfo.ulFlags + BIF_BROWSEINCLUDEFILES

    If (ShowNewFolderButton Or NewStyle) And isME2KXP Then
        tBrowseInfo.ulFlags = tBrowseInfo.ulFlags Or BIF_NEWDIALOGSTYLE + BIF_UAHINT
        'OleInitialize Null ' Initialize OLE and COM
    Else
        tBrowseInfo.ulFlags = tBrowseInfo.ulFlags Or BIF_STATUSTEXT
    End If
    If Not ShowNewFolderButton Then tBrowseInfo.ulFlags = tBrowseInfo.ulFlags Or BIF_NONEWFOLDERBUTTON
    tBrowseInfo.ulFlags = tBrowseInfo.ulFlags Or Flags
    tBrowseInfo.hOwner = OwnerForm
    lpIDList = API_shellMethods_SHBrowseForFolder(tBrowseInfo)
    If clRoot = True Then Call API_shellMethods_CoTaskMemFree(tBrowseInfo.pidlRoot)

    If (lpIDList) Then
        sBuffer = Space$(MAX_PATH)
        Call API_shellMethods_SHGetPathFromIDList(lpIDList, sBuffer)
        Call API_shellMethods_CoTaskMemFree(lpIDList)
        sBuffer = Left$(sBuffer, InStr(sBuffer, vbNullChar) - 1)
        Dialogs_Browse = sBuffer
    Else
        throw OperationCanceledException
    End If
End Function

Private Function GetAddressOfFunction(zAdd As Long) As Long
    GetAddressOfFunction = zAdd
End Function

Public Function GetSpecialFolderID(ByVal Folder As API_SpecialFolders) As Long
    Dim IDL As ITEMIDLIST, R As Long
    R = API_shellMethods_SHGetSpecialFolderLocation(ByVal 0&, Folder, IDL)


    If R = 0 Then
        GetSpecialFolderID = IDL.mkid.cb
    Else
        GetSpecialFolderID = 0
    End If
End Function

Private Function Browse_CallbackProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal lp As Long, ByVal pData As Long) As Long
    On Local Error Resume Next
    Dim sBuffer As String

    Select Case uMsg
        Case BFFM_INITIALIZED
        Call API_shellMethods_SendMessage(hWnd, BFFM_SETSELECTION, 1, m_CurrentDirectory)
        If OK_BUTTON_TEXT <> vbNullString Then Call API_shellMethods_SendMessage2(hWnd, BFFM_SETOKTEXT, 1, StrPtr(OK_BUTTON_TEXT))
        Case BFFM_SELCHANGED
        sBuffer = Space$(MAX_PATH)
        Call API_shellMethods_SHGetPathFromIDList(lp, sBuffer)
        sBuffer = Left$(sBuffer, InStr(sBuffer, vbNullChar) - 1)


        If Len(sBuffer) = 0 Then
            Call API_shellMethods_SendMessage2(hWnd, BFFM_ENABLEOK, 1, 0)
            Call API_shellMethods_SendMessage(hWnd, BFFM_SETSTATUSTEXT, 1, "")
        Else
            Call API_shellMethods_SendMessage(hWnd, BFFM_SETSTATUSTEXT, 1, sBuffer)
        End If
    End Select
    Browse_CallbackProc = 0
End Function



Public Sub OpenURL(ByVal hWnd As Long, ByVal URL As String)
    Call API_shellMethods_ShellExecute(hWnd, "open", URL, vbNullString, vbNullString, vbNormal)
End Sub
