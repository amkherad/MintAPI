VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Simple Encryption."
   ClientHeight    =   4020
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4740
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4020
   ScaleWidth      =   4740
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Caption         =   "Shell Extensions"
      Height          =   975
      Left            =   120
      TabIndex        =   7
      Top             =   1920
      Width           =   4575
      Begin VB.CommandButton Command5 
         Caption         =   "Unregister Shell Extensions"
         Height          =   255
         Left            =   443
         TabIndex        =   9
         Top             =   600
         Width           =   3855
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Register Shell Extensions"
         Height          =   255
         Left            =   443
         TabIndex        =   8
         Top             =   240
         Width           =   3855
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Encryption."
      Height          =   1695
      Left            =   105
      TabIndex        =   2
      Top             =   120
      Width           =   4575
      Begin VB.CommandButton Command4 
         Cancel          =   -1  'True
         Caption         =   "Cancel"
         Enabled         =   0   'False
         Height          =   375
         Left            =   3360
         TabIndex        =   11
         Top             =   1200
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Encrypt"
         Default         =   -1  'True
         Enabled         =   0   'False
         Height          =   375
         Left            =   2160
         TabIndex        =   10
         Top             =   1200
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Delete original file."
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   1200
         Width           =   1695
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   240
         TabIndex        =   5
         Text            =   "[Password..]"
         Top             =   720
         Width           =   2655
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   240
         OLEDropMode     =   1  'Manual
         TabIndex        =   4
         Top             =   240
         Width           =   2655
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Browse"
         Height          =   375
         Left            =   3360
         TabIndex        =   3
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Instructions"
      Height          =   975
      Left            =   105
      TabIndex        =   0
      Top             =   3000
      Width           =   4575
      Begin VB.Label Label1 
         Caption         =   "First click the Browse button and select the file that you want to encrypt/decrypt."
         Height          =   495
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   4215
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Text
'Test project. A simple file encryption utility
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Sub CopyMemory Lib "KERNEL32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function IsUserAnAdmin Lib "shell32.dll" () As Long

Private Type OPENFILENAME
        lStructSize As Long
        hwndOwner As Long
        hInstance As Long
        lpstrFilter As String
        lpstrCustomFilter As String
        nMaxCustFilter As Long
        nFilterIndex As Long
        lpstrFile As Long
        nMaxFile As Long
        lpstrFileTitle As String
        nMaxFileTitle As Long
        lpstrInitialDir As String
        lpstrTitle As String
        flags As Long
        nFileOffset As Integer
        nFileExtension As Integer
        lpstrDefExt As String
        lCustData As Long
        lpfnHook As Long
        lpTemplateName As String
End Type

Private Declare Function CommDlgExtendedError Lib "comdlg32.dll" () As Long
Private Declare Function GetLongPathName Lib "KERNEL32" Alias "GetLongPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long

Private lastOutputFile As String, CancelOperation As Boolean, Password As String
Private Buffer() As Byte
Private WithEvents EncryptionThread As Thread
Attribute EncryptionThread.VB_VarHelpID = -1


Private Sub Command1_Click()
Dim x As OPENFILENAME
On Error GoTo errhnd:
Debug.Print Len(x)
Dim filename As String
ResetAll
filename = vbNullChar & Space$(254)
StrConv filename, vbFromUnicode
MsgBox Me.hwnd
x.hwndOwner = Me.hwnd
x.lStructSize = Len(x)
x.nMaxFile = 256
x.lpstrFile = StrPtr(filename)

Debug.Print GetOpenFileName(x)
filename = StrConv(filename, vbUnicode)
Text1.Text = filename
Label1.Caption = "Now Enter a 5 or more character long password."
Text2.SetFocus
errhnd:
End Sub




Private Sub Command2_Click()
Dim outFile As String
If Text1.Text = "" Then Exit Sub

If Right(Text1.Text, 3) = "enc" Then outFile = Left(Text1.Text, Len(Text1.Text) - 4) _
                            Else outFile = Text1.Text & ".enc"
CheckFile outFile

MsgBox "Now going to encode the file. The file will be saved an extension "".enc"" in the same directory", vbOK Or vbInformation
Command4.Enabled = True: Command2.Enabled = False
'EncryptionThread_DoWork Text1.Text & vbCrLf & outFile, 0
EncryptionThread.StartWorkAsync Me, Text1.Text & vbCrLf & outFile
'ResetAll
'Command4.Enabled = False: Command2.Enabled = True

Exit Sub
If EncodeAndDecode(Text1.Text, outFile) = 0 Then
  MsgBox "File Encoded successfully", vbOK Or vbInformation
Else
  MsgBox "File Encoding Unsuccessful. Error occured or user interrupted.", vbOK Or vbCritical
End If
ResetAll
Command4.Enabled = False: Command2.Enabled = True

End Sub


Private Sub Command3_Click()
If IsUserAnAdmin = 1 Then
modShellExtension.SetShellExtension
MsgBox "Application registered in the shell. Now you can simply right click any file to encrypt/decrypt it.", vbOKOnly Or vbInformation
Else
MsgBox "Sorry. Only Administrators can Register or Unregister the shell extensions. Run this application as administrator or tell your administrator to do so." _
            & vbCrLf & "To run this app as administrator. Right click and select Run As..", vbOKOnly Or vbExclamation
End If
Form1.SetFocus
End Sub

Private Sub Command4_Click()
'CancelOperation = True
EncryptionThread.CancelWork
End Sub

Private Sub Command5_Click()
If IsUserAnAdmin = 1 Then
modShellExtension.RemoveShellExtension
MsgBox "Application unregistered. Click ""Register Shell Extensions"" to re-register it.", vbOKOnly Or vbInformation
Else
MsgBox "Sorry. Only Administrators can Register or Unregister the shell extensions. Run this application as administrator or tell your administrator to do so." _
            & vbCrLf & "To run this app as administrator. Right click and select Run As..", vbOKOnly Or vbExclamation
End If
Form1.SetFocus
End Sub

Private Sub EncryptionThread_DoWork(ByVal Arg As Variant, Result As Variant)
Dim srcFile As String, outFile As String
srcFile = Left(CStr(Arg), InStr(1, CStr(Arg), vbCrLf) - 1)
outFile = Right(CStr(Arg), Len(CStr(Arg)) - Len(srcFile) - 2)
Result = EncodeAndDecode(srcFile, outFile)
End Sub

Private Sub EncryptionThread_ProgressChanged(ByVal ProgressPercent As Long)
Static Time As Single
If Time = 0 Then Time = Timer
Label1.Caption = "Encryption in progress. Completed = " & ProgressPercent & "%" & _
                    vbCrLf & "Elapsed time = " & (Timer - Time) & " seconds."

End Sub

Private Sub EncryptionThread_WorkCompleted(ByVal Result As Variant)
If Result = 0 Then
  MsgBox "File Encoded successfully", vbOK Or vbInformation
Else
  MsgBox "File Encoding Unsuccessful. Error occured or user interrupted.", vbOK Or vbCritical
End If
ResetAll
Command4.Enabled = False: Command2.Enabled = True
End Sub

Private Sub Form_Activate()
If Command$ <> "" Then
    Text1.Text = GetLongName(Trim$(Command))
    If Right(Text1.Text, 3) = "enc" Then Check1.Value = 1
    Text1_Validate False
    Text2.SetFocus
End If
End Sub

Private Sub Form_Load()
ChDir App.Path
Text1.Text = ""
Set EncryptionThread = New Thread
End Sub

Public Function EncodeAndDecode(ByVal inpFile As String, ByVal outFile As String) As Long
Dim txt As String, i As Long, Tmp As Single, Time As Single, Size As Long ', pcent As Single
Dim pcentex As Single, j As Long, BytesToRW As Long

On Error GoTo errHandler:

'Timer to calculate the seconds taken
'Time = Timer

'First generate XORMask
txt = Password

Rnd (-1)
Randomize Len(txt)

For i = 1 To Len(txt)
 Sum = Sum + Asc(Mid(txt, i, 1))
Next

For i = 1 To Sum
Tmp = Rnd()
Next

Dim XorMask As Integer
XorMask = Int(Rnd * 256 + 1)


'XORMask done. Now load input file and create output file.

Dim F1 As Long, F2 As Long, char As Byte
F1 = FreeFile: F2 = F1 + 1

Open inpFile For Binary As F1
Open outFile For Binary As F2

Size = FileLen(inpFile)

BytesToRW = 1024 * 1 - 1
ReDim Buffer(BytesToRW)

'Pcent to track the progress of encryption
pcentex = (CSng(Size) / BytesToRW) / 100
If pcentex < 0 Then pcentex = 1
Frame1.Caption = "Status"


'We shall operate 1 KB at one time. So jump for size\1024 times. Remaining bytes
'shall be processed at the end.
For i = 1 To Size \ (BytesToRW + 1)
    
Get F1, , Buffer
    
    For j = 0 To BytesToRW
     Buffer(j) = Buffer(j) Xor XorMask
    Next

If i Mod pcentex = 0 Then
    'Label1.Caption = "Encryption in progress. Completed = " & i \ pcentex & "%" & _
    '                    vbCrLf & "Elapsed time = " & (Timer - Time) & " seconds."
    EncryptionThread.ReportProgress i \ pcentex
    'DoEvents
End If

If EncryptionThread.CancellationPending = False Then
    'CancelOperation = False
    GoTo Abort:
End If

ResumeNow:
Put F2, , Buffer
Next

'Now do the remaining bytes.
Dim BufferEx() As Byte
Tmp = Size - Loc(2) - 1
ReDim BufferEx(Tmp)


Get F1, , BufferEx

For j = 0 To Tmp
    BufferEx(j) = BufferEx(j) Xor XorMask
    Next
Put F2, , BufferEx

'Encrypted successfully. Now close files calculate time and set the label.
Close F1, F2
Time = Timer - Time
Label1.Caption = "Last encrypted file was " & Str((FileLen(inpFile) \ 1024)) & " KB long." & vbCrLf & _
                    "And last encryption operation took " & Fix(Time) & " seconds."
lastOutputFile = outFile
If Check1.Value = 1 Then Kill inpFile

EncodeAndDecode = 0

Exit Function

Abort:
'If MsgBox("Do you really want to Abort?", vbYesNo Or vbExclamation, "Abort?") = vbYes Then
    Close F1, F2
    Kill outFile
    EncodeAndDecode = -1
    Exit Function
'Else
'    GoTo ResumeNow
'End If

errHandler:
EncodeAndDecode = err.Number
Debug.Print err.Description
Resume Next
End Function


Private Sub Form_Unload(Cancel As Integer)
Set EncryptionThread = Nothing
End Sub

Private Sub Text1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, Y As Single)
On Error GoTo err:
'Debug.Print Data.GetData(20)
Text1.Text = Data.Files.Item(1)
If Right(Text1.Text, 3) = "enc" Then Check1.Value = 1
Text1_Validate False
err:

End Sub

Private Sub Text1_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, Y As Single, State As Integer)
Dim c As OLEDropEffectConstants
If Data.GetFormat(vbCFFiles) Then Effect = vbDropEffectCopy Else Effect = vbDropEffectNone
If Data.Files.Count > 1 Then Effect = vbDropEffectNone
End Sub

Private Sub Text1_Validate(Cancel As Boolean)
On Error GoTo error:

If Text1.Text = "" Then Exit Sub
FileLen Text1.Text
Label1.Caption = "Now Enter a 5 or more character long password."
Text2.SetFocus
Cancel = False
Exit Sub

error:
MsgBox "The specified file doesn't exists. Please re-enter", vbOKOnly Or vbExclamation
Cancel = True
End Sub

Private Sub Text2_Change()
If Len(Text2.Text) >= 5 Then Command2.Enabled = True Else Command2.Enabled = False
End Sub

Private Sub Text2_GotFocus()
If Text1.Text = "" Then
Text1.SetFocus
Exit Sub
End If
Text2.Text = ""
Text2.PasswordChar = "*"
End Sub

Private Sub Text2_Validate(Cancel As Boolean)
If Len(Text2.Text) < 5 Then
    MsgBox "Password must be greater or equal to 5 characters. Please re-enter the password", vbExclamation Or vbOKOnly
    Text2.Text = ""
    Cancel = True
Else
    Password = Text2.Text
    Text2.Text = "[Password..]"
    Text2.PasswordChar = ""
    Command2.Enabled = True
    Label1.Caption = "Now click the Encrypt Button to start encryption."
    Cancel = False
End If
End Sub

Private Sub ResetAll()
Text1.Text = ""
Text2.Text = "[Password..]"
Text2.PasswordChar = ""
Text1.SetFocus
Command2.Enabled = False
Command4.Enabled = False
Check1.Value = 0
Label1.Caption = "First click the Browse button and select the file that you want to encrypt/decrypt."
Frame1.Caption = "Instruction"
CancelOperation = False
End Sub

Private Sub CheckFile(File As String)
On Error GoTo error:
FileLen File
File = GetFilenameWithoutExtension(File) & "_" & File
CheckFile File
error:
End Sub

Private Function GetFilenameWithoutExtension(File As String) As String
Dim cnt As Integer, Tmp As String
cnt = 1
While cnt <> 0
    Tmp = Left(File, cnt)
    cnt = InStr(cnt + 1, File, "\")
Wend
cnt = InStr(Len(Tmp), File, ".")
cnt = cnt - 1
GetFilenameWithoutExtension = Left(File, cnt)
File = Right(File, Len(File) - cnt)
End Function

Private Function GetLongName(FilePath As String) As String
Dim ASD As String * 225
GetLongPathName FilePath, ASD, 225
GetLongName = ASD
End Function
