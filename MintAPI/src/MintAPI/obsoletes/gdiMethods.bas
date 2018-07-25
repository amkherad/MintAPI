Attribute VB_Name = "gdiMethods"
'@PROJECT_LICENSE
Option Explicit
Option Base 0
Const CLASSID As String = "gdiMethods"

Private Const DM_BITSPERPEL = &H40000
Private Const DM_PELSWIDTH = &H80000
Private Const DM_PELSHEIGHT = &H100000
Private Const CDS_UPDATEREGISTRY = &H1
Private Const CDS_TEST = &H4
Private Const WM_DISPLAYCHANGE = &H7E
Private Const HWND_BROADCAST = &HFFFF&
Private Const CCHDEVICENAME = 32
Private Const CCHFORMNAME = 32
Private Const DISP_CHANGE_SUCCESSFUL = 0
Private Const DISP_CHANGE_RESTART = 1


Private Type DEVMODE
    dmDeviceName As String * CCHDEVICENAME
    dmSpecVersion As Integer
    dmDriverVersion As Integer
    dmSize As Integer
    dmDriverExtra As Integer
    dmFields As Long
    dmOrientation As Integer
    dmPaperSize As Integer
    dmPaperLength As Integer
    dmPaperWidth As Integer
    dmScale As Integer
    dmCopies As Integer
    dmDefaultSource As Integer
    dmPrintQuality As Integer
    dmColor As Integer
    dmDuplex As Integer
    dmYResolution As Integer
    dmTTOption As Integer
    dmCollate As Integer
    dmFormName As String * CCHFORMNAME
    dmLogPizels As Integer
    dmBitsPerPel As Long
    dmPelsWidth As Long
    dmPelsHeight As Long
    dmDisplayFlags As Long
    dmDisplayFrequency As Long
    dmICMMethod As Long
    dmICMIntent As Long
    dmMediaType As Long
    dmDitherType As Long
    dmReserved1 As Long
    dmReserved2 As Long
End Type


Private Declare Function API_gdiMethods_SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function API_gdiMethods_EnumDisplaySettings Lib "user32" Alias "EnumDisplaySettingsA" (ByVal lpszDeviceName As Long, ByVal iModeNum As Long, lpDevMode As Any) As Boolean
Private Declare Function API_gdiMethods_ChangeDisplaySettings Lib "user32" Alias "ChangeDisplaySettingsA" (lpDevMode As Any, ByVal dwFlags As Long) As Long
Private Declare Function API_gdiMethods_GetDeviceCaps Lib "user32" Alias "GetDeviceCaps" (ByVal hdc As Long, ByVal nIndex As Long) As Long

Dim inited As Boolean

Public Sub Initialize()
    If inited Then Exit Sub
    inited = True
End Sub
Public Sub Dispose(Optional ByVal Force As Boolean = False)
    If Not inited Then Exit Sub
    inited = False
End Sub



Public Function ChangeResolution(X As Long, Y As Long, Bits As Long, Optional ByVal SendMessageToWindows As Boolean = True) As Boolean
    Dim DevM As DEVMODE, erg As Long
    Dim ScInfo As Long
    'Get the info into DevM
    erg = API_gdiMethods_EnumDisplaySettings(0&, 0&, DevM)
    'This is what we're going to change
    DevM.dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT Or DM_BITSPERPEL
    DevM.dmPelsWidth = X 'ScreenWidth
    DevM.dmPelsHeight = Y 'ScreenHeight
    DevM.dmBitsPerPel = Bits '(can be 8, 16, 24, 32 or even 4)
    'Now change the display and check if possible
    erg = API_gdiMethods_ChangeDisplaySettings(DevM, CDS_TEST)
    'Check if succesfull
    Select Case erg&
        Case DISP_CHANGE_RESTART
            ChangeResolution = False
        Case DISP_CHANGE_SUCCESSFUL
            erg = API_gdiMethods_ChangeDisplaySettings(DevM, CDS_UPDATEREGISTRY)
            ScInfo = ((Y * (2 ^ 16)) + X)
            'Notify all the windows of the screen resolution change
            If SendMessageToWindows Then Call API_gdiMethods_SendMessage(HWND_BROADCAST, WM_DISPLAYCHANGE, ByVal Bits, ByVal ScInfo)
            ChangeResolution = True
        Case Else
            ChangeResolution = False
    End Select
End Function
