Attribute VB_Name = "modMain"
Option Explicit

Public Sub Main()
    Dim args() As String
    If Not (InStr(1, LCase(Command), "noremain") > 0) Then Call Load(mFrm)
    Dim miccyUltimateTools As Object
    On Error GoTo err
    Set miccyUltimateTools = CreateObject("miccyUltimateTools.API")
err:
    If miccyUltimateTools Is Nothing Then
        
        Exit Sub
    End If
    Call miccyUltimateTools.Initialize(mFrm)
    If miccyUltimateTools.CanLoadNewInstance Then
        Call miccyUltimateTools.i_Main(args())
    Else
        Call miccyUltimateTools.clog("Can't Start A New Instance.")
        Call MsgBox("Can't Start A New Instance.")
        Call Unload(mFrm)
        End
    End If
End Sub

Public Sub RegisterActiveX(ByVal Path)
    
End Sub
