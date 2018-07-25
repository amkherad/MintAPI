VERSION 1.0 CLASS
BEGIN
  MultiUse = 0   'False
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsNewSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Event Destroy()

Private WithEvents mobjGateway  As Gateway
Attribute mobjGateway.VB_VarHelpID = -1

Private Sub Class_Initialize()
    Set mobjGateway = New Gateway
    mobjGateway.HandleAsyncErrors = True
End Sub

Private Sub Class_Terminate()
    Set mobjGateway = Nothing
End Sub

Public Property Get hWnd() As Long
    hWnd = mobjGateway.hWnd
End Property

Public Sub Link(ByVal hWnd As Long)
    mobjGateway.StartLink hWnd
End Sub

Private Sub mobjGateway_AsynchronousError(ByVal Number As Long, ByVal Source As String, ByVal Description As String)
    MsgBox Err.Description, vbCritical, "Thread Error"
End Sub

Private Sub mobjGateway_DataArrived(Data As Variant, ByVal Synchronous As Boolean)
Dim frmNew  As frmSearch
Dim objPB   As PropertyBag
On Error GoTo ErrHandler
    Set objPB = New PropertyBag
    objPB.Contents = Data
    Set frmNew = New frmSearch
    Set frmNew.NewSearch = Me
    Load frmNew
    frmNew.Show
    frmNew.StartSearch objPB.ReadProperty("PATH"), objPB.ReadProperty("EXTENSION")
    Set frmNew = Nothing
    Exit Sub
ErrHandler:
    mobjGateway.RaiseError Err.Number, Err.Source, Err.Description
End Sub

Private Sub mobjGateway_LinkTerminated()
    RaiseEvent Destroy
End Sub

Friend Sub FoundFile(ByVal File As String, ByVal ReadOnly As Boolean, ByVal Hidden As Boolean)
Static lngCount As Long
Dim objPB       As PropertyBag
Dim bytData()   As Byte
    lngCount = lngCount + 1
    If mobjGateway.StackCount < 10 Then
        Set objPB = New PropertyBag
        objPB.WriteProperty "HWND", mobjGateway.hWnd
        objPB.WriteProperty "COUNT", lngCount
        objPB.WriteProperty "FILE", File
        objPB.WriteProperty "READONLY", ReadOnly
        objPB.WriteProperty "HIDDEN", Hidden
        bytData = objPB.Contents
        Set objPB = Nothing
        mobjGateway.PostData bytData
    End If
    Sleep 1
End Sub
