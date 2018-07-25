Attribute VB_Name = "imaging"
Option Explicit
Public Sub StartImaging(Path As String, SaveTo As String)
    Dim f As File
    Set f = File(SaveTo)
    Call f.Create(fmCreate, fNormal, faWrite, fshNone)

On Error GoTo err

    Call f.WriteLine("@ Directory Imager 0001")
    Call f.WriteLine("Path: """ & Path & """")
    Call f.WriteLine("-----------------------")

    Call recursiveImageDir(Path, f)

err:
    Call f.CloseFile
End Sub
Private Sub recursiveImageDir(Path As String, f As File)
    Dim dirs() As String, Files() As String
    Dim i As Long, dr As Directory

    Set dr = Directory(Path)
    dirs = dr.SubDirectories(vbDirectory, True)

    frmMain.log.Caption = "" & dr.AbsolutePath

    On Error Resume Next
    Call f.WriteLine("Path: """ & dr.Name & """")
    Files = dr.FileNames(vbNormal, False)
    For i = 0 To ArraySize(Files) - 1
        Call f.WriteLine(Files(i))
    Next

    For i = 0 To ArraySize(dirs) - 1
        Call recursiveImageDir(dirs(i), f)
        Call f.WriteLine("Path: ""..""")
    Next
End Sub


'Public Sub StartImaging(Path As String, SaveTo As String)
'    Dim f As Long
'    f = FreeFile
'    Open SaveTo For Output As #f
'
'On Error GoTo err
'
'    Print #f, ("@ Directory Imager 0001")
'    Print #f, ("Path: """ & Path & """")
'    Print #f, ("-----------------------")
'
'    Call recursiveImageDir(Path, f)
'
'err:
'    Close #f
'End Sub
'Private Sub recursiveImageDir(Path As String, f As Long)
'    Dim dirs() As String, Files() As String
'    Dim i As Long, dr As Directory
'
'    Set dr = Directory(Path)
'    dirs = dr.SubDirectories(vbDirectory, True)
'
'    frmMain.Log.Caption = "" & dr.AbsolutePath
'
'    On Error Resume Next
'    Print #f, ("Path: """ & dr.Name & """")
'    Files = dr.FileNames(vbNormal, False)
'    For i = 0 To ArraySize(Files) - 1
'        Print #f, (Files(i))
'    Next
'
'    For i = 0 To ArraySize(dirs) - 1
'        Call recursiveImageDir(dirs(i), f)
'        Print #f, ("Path: ""..""")
'    Next
'End Sub
