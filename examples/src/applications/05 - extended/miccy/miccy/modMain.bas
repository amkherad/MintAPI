Attribute VB_Name = "modMain"
Option Explicit

Public Declare Function API_CreateRoundRectRgn Lib "gdi32" Alias "CreateRoundRectRgn" (ByVal x1 As Integer, ByVal y1 As Integer, ByVal x2 As Integer, ByVal y2 As Integer, ByVal x3 As Integer, ByVal y3 As Integer) As Long
Public Declare Function API_DeleteObject Lib "gdi32" Alias "DeleteObject" (ByVal hObject As Long) As Long
Public Declare Function API_SetWindowRgn Lib "user32" Alias "SetWindowRgn" (ByVal hwnd As Long, ByVal hRGN As Long, ByVal bRedraw As Boolean) As Long
Public Declare Function API_SetParent Lib "user32" Alias "SetParent" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long

Public Iwindows As Collection

Public actionBuf As ActionBuffer
Public lastProc As Process
Public provider As provider

Public callerLibrary_unloadObject As Object

Private log_path As String
Private log_state As Boolean

Public Sub Main()
    clog "Initializing miccy..."
'    Call baseConstants.Initialize
'    Call Exceptions.Initialize
'    Call kernelMethods.Initialize
'    Call baseMethods.Initialize
'    Call baseMethods2.Initialize
'    Call bitOperations.Initialize
log_state = True
    If Not checklogfile Then
        
    End If
    Call ReadConfig
    Call CloseConfig
End Sub
Public Sub EndApp()
    clog "Unloading objects..."
    clog "Terminating application..."
    Call CloseConfig
    Call FinilizeAllPlugins
    On Error Resume Next
    If Not callerLibrary_unloadObject Is Nothing Then Call Unload(callerLibrary_unloadObject)
    Call Unload(Settings)
    Call Unload(mForm)
    Call Unload(About)
    Call Unload(toolsWindow)
    clog "Application successfully terminated."
    clog "Disposing miccy..."
log_state = False
    clog
End Sub

Public Sub FinilizeAllPlugins()
    Call SaveConfig
End Sub

Private Function checklogfile() As Boolean
'On Error GoTo err
    log_path = Directory.ConcatPath("", "miccy.log.html")
    If Not File(log_path).Exists Then
        Dim fl As Long
        Dim pa As String
        fl = FreeFile
        pa = File(log_path).Location
        Dim d As Directory
        Set d = Directory(pa)
        If Not d.Exists Then Call d.Create
        Open log_path For Binary As #fl
        Close #fl
        Dim s As String
        s = _
        "<html><head>" & _
        "<style rel=""stylesheet"" type=""text/css"">" & _
        " *{font-size:14px;font-family:Consolas,Courier,Sans Serif;color:#333;}" & _
        " body{background-color:#b7c6cc;font-family:Consolas,Courier,Sans Serif;}" & _
        " a , a:visited{font-family:Consolas;font-size:14px;color:#194954;text-decoration:none;}" & _
        " a:hover{text-decoration:underline;}" & _
        " h1{margin-right:100px;font-size:17px;color:#013a62;vertical-align:bottom;}" & _
        " h5{margin-right:120px;font-size:9px;color:#013a62;}" & _
        " .dt{color:green;}" & _
        " .un{color:blue;}" & _
        " .ln{height:9px;background-color:#eee;border:none;border-bottom:1px solid #ddd;}" & _
        " .dm{color:6a018f;}" & _
        " .err{color:red;text-weight:bold;}" & _
        " #content{background-color:#eee;direction:ltr;margin:0px 80px 0px 80px;border:1px solid #aaa;padding:30px;}" & _
        " #ulist{background-color:#eee;padding:1;border:1px solid #ddd;list-style:none;}" & _
        " #ulist li.lit{background-color:#fff;padding:0px 10px 0px 10px;border-bottom:1px solid #ddd;line-height:30px;}" & _
        " #ulist li.lit:hover{background-color:#fafafa;}" & _
        "</style>" & _
        "</head><body><h1>miccy Ultimate Tools</h1> <br><h5>programmer: Ali Mousavi Kherad | <a href=""mailto:alimousavikherad@gmail.com"">alimousavikherad@gmail.com</a></h5>" & _
        "<div id=""content""><ul id=""ulist"">" + vbCrLf
        Call print_log(s)
    End If
    checklogfile = True
err:
    checklogfile = False
End Function
Private Sub print_log(ByRef msg As String)
    Dim fl As Long
    On Error GoTo err
    fl = FreeFile
    Open log_path For Append As #fl
    On Error GoTo err1
    '------------------------------
    Print #fl, msg
    '------------------------------
err1:
    Close #fl
err:
End Sub
Public Sub clog(Optional ByVal Message As String, Optional ByVal flags As Long = &H1)
    If (flags And &H1) = &H1 Then Debug.Print "application log:" & Message
    Dim s As String
    If Not log_state Then
        s = "<li class=""ln""></li>" & vbCrLf
    Else
        s = _
            "<li class=""lit""><span class=""dt"">" & _
               Replace(Format(Now, ""), "  ", " &nbsp;") & "</span> - " & _
               Message & _
            "</li>" & vbCrLf
    End If
    Call print_log(s)
End Sub
Public Sub cdebug(ByVal Message As String)
    Debug.Print "application debug-message:" & Message
    Call clog("<span class=""dm"">debug-message:</span><span class=""err"">" & Message & "</span>", 0)
End Sub
Public Sub cerr(ByVal Message As String)
    Call cdebug(Message)
    throw Exps.Exception(Message)
End Sub
Public Sub cfatal(ByVal Message As String)
    Call cdebug(Message)
    Call modMain.EndApp
End Sub
Public Sub cthrow(ByVal Message As String)
    Call cerr(Message)
End Sub


Public Sub ClosePlugins()
    
End Sub
Public Sub OpenPlugin(uniqueID As String)
    
End Sub
Public Sub ClosePlugin(uniqueID As String)
    
End Sub

Public Function CollectPluginsInformation() As IPlugin()
    
End Function

Public Sub regActive(ByVal Path As String)
    
End Sub

Public Function installPlugin(ByVal Path As String)
    
End Function
Public Sub uninstallPlugin()
    
End Sub
Public Function InstallAction()
    
End Function
Public Sub UninstallAction()
    
End Sub
Public Function InstallFilter()
    
End Function
Public Sub UninstallFilter()
    
End Sub

Public Sub appendWindow(ByVal win As Iwindow)
    
End Sub
Public Sub removeWindow(ByVal win As Iwindow)
    
End Sub

Public Sub ShowWindow(ByVal win As Iwindow)
    
End Sub
Public Sub HideWindow(ByVal win As Iwindow)
    
End Sub
Public Sub tempshowWin(ByVal win As Iwindow)
    
End Sub

Private Sub CopyUniqueID(import As String, importLen As Long, export As String, useNulls As Boolean)
    export = ""
    Dim i As Long
    If importLen <= 0 Then Exit Sub
    If useNulls Then
        For i = 1 To importLen
            export = export & Mid(import, i, 1)
        Next
    Else
        Dim ch As String * 1
        For i = 1 To importLen
            ch = Mid(import, i, 1)
            If ch <> vbNullChar Then export = export & ch
        Next
    End If
End Sub
'
''configFL
'Private Sub WriteRecordArray(arr() As GLBTP, res As Long)
'    Dim Count As Long
'    Get #configFL, res, Count
'    Erase arr()
'    If Count = 0 Then Exit Sub
'    If Count < 0 Then
'        Put #configFL, res, 0
'    End If
'End Sub
'Private Sub ReadRecordArray(arr() As GLBTP, res As Long)
'
'End Sub
