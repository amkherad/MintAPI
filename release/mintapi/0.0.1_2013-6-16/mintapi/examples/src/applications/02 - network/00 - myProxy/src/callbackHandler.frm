VERSION 5.00
Begin VB.Form callbackHandler 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu m 
      Caption         =   "Tray Menu"
      Begin VB.Menu m_stablish 
         Caption         =   "&Stablish Proxy"
      End
      Begin VB.Menu m_settings 
         Caption         =   "&Settings..."
      End
      Begin VB.Menu m_exit 
         Caption         =   "&Exit"
      End
   End
End
Attribute VB_Name = "callbackHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function API_FlashWindow Lib "user32" Alias "FlashWindow" (ByVal hwnd As Long, ByVal bInvert As Long) As Long
Private Declare Function API_BringWindowToTop Lib "user32" Alias "BringWindowToTop" (ByVal hwnd As Long) As Long

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If X = 7725 Then frmMain.Visible = Not frmMain.Visible
    If X = 7755 Then Call PopupMenu(M)
    If X = PREVINSTANCEMESSAGEID_RECIEVE Then 'if previnstance run then run a new instance of application.
        If Not frmMain.Visible Then Call frmMain.Show 'show main window.
        Call API_BringWindowToTop(frmMain.hwnd)
        Call API_FlashWindow(frmMain.hwnd, 5) 'flashing window after showing it.
    End If
End Sub

Private Sub m_exit_Click()
    Call ExitApplication
End Sub

Private Sub mt_Show_Click()
    frmMain.Visible = Not frmMain.Visible
End Sub

