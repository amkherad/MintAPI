Attribute VB_Name = "modShellExtension"
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, phkResult As Long, lpdwDisposition As Long) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Private Declare Function RegFlushKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.
Private Const HKEY_CLASSES_ROOT = &H80000000
Private Const KEY_CREATE_SUB_KEY = &H4
Private Const KEY_SET_VALUE = &H2

Private Type SECURITY_ATTRIBUTES
        nLength As Long
        lpSecurityDescriptor As Long
        bInheritHandle As Long
End Type

'Some variables
Private Hnd_AllFile As Long, Hnd_EncApp As Long, Hnd_EncFile As Long


Public Sub SetShellExtension()
Dim AppNAme As String, Y As SECURITY_ATTRIBUTES, Appicon As String
AppNAme = App.Path & "\" & App.EXEName & ".exe %1"
Appicon = App.Path & "\icon.ico"

'Do the all files
Debug.Print RegCreateKeyEx(HKEY_CLASSES_ROOT, "*\shell\Encrypt/Decrypt\command", 0, "", 0, KEY_SET_VALUE, Y, Hnd_AllFile, 0)
Debug.Print RegSetValueEx(Hnd_AllFile, "", 0, 1, ByVal AppNAme, Len(AppNAme) + 1)
Debug.Print RegCloseKey(Hnd_AllFile)

'do the .enc section
Debug.Print RegCreateKeyEx(HKEY_CLASSES_ROOT, ".enc", 0, "", 0, KEY_SET_VALUE, Y, Hnd_EncFile, 0)
Debug.Print RegSetValueEx(Hnd_EncFile, "", 0, 1, ByVal "encfile", 8)
Debug.Print RegCloseKey(Hnd_EncFile)

'do the encfile section
Debug.Print RegCreateKeyEx(HKEY_CLASSES_ROOT, "encfile\shell\Encrypt/Decrypt\command", 0, "", 0, KEY_SET_VALUE, Y, Hnd_EncApp, 0)
Debug.Print RegSetValueEx(Hnd_EncApp, "", 0, 1, ByVal AppNAme, Len(AppNAme) + 1)
Debug.Print RegCloseKey(Hnd_EncApp)


Debug.Print RegCreateKeyEx(HKEY_CLASSES_ROOT, "encfile\defaulticon", 0, "", 0, KEY_SET_VALUE, Y, Hnd_EncApp, 0)
Debug.Print RegSetValueEx(Hnd_EncApp, "", 0, 1, ByVal Appicon, Len(Appicon) + 1)
Debug.Print RegCloseKey(Hnd_EncApp)
End Sub

Public Sub RemoveShellExtension()
Debug.Print RegDeleteKey(HKEY_CLASSES_ROOT, "*\shell\Encrypt/Decrypt\command")
Debug.Print RegDeleteKey(HKEY_CLASSES_ROOT, "*\shell\Encrypt/Decrypt")

Debug.Print RegDeleteKey(HKEY_CLASSES_ROOT, ".enc")

Debug.Print RegDeleteKey(HKEY_CLASSES_ROOT, "encfile\shell\Encrypt/Decrypt\command")
Debug.Print RegDeleteKey(HKEY_CLASSES_ROOT, "encfile\shell\Encrypt/Decrypt")
Debug.Print RegDeleteKey(HKEY_CLASSES_ROOT, "encfile\shell")
Debug.Print RegDeleteKey(HKEY_CLASSES_ROOT, "encfile\defaulticon")
Debug.Print RegDeleteKey(HKEY_CLASSES_ROOT, "encfile")


End Sub

