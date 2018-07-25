VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)

Private WithEvents mobjGateway  As Gateway
Attribute mobjGateway.VB_VarHelpID = -1

Private mlngParentPtr   As Long

Private Sub Class_Initialize()
    Set mobjGateway = New Gateway
End Sub

Private Sub Class_Terminate()
    Set mobjGateway = Nothing
End Sub

Public Property Get Searches() As clsSearches
    Set Searches = PtrObj(mlngParentPtr)
End Property

Private Function PtrObj(ByVal Pointer As Long) As Object
Dim objObject   As Object
    CopyMemory objObject, Pointer, 4&
    Set PtrObj = objObject
    CopyMemory objObject, 0&, 4&
End Function

Public Property Get Key() As String
    Key = "ID" & mobjGateway.hWnd
End Property

Public Sub Initialise(ByRef pobjSearches As clsSearches)
    mlngParentPtr = ObjPtr(pobjSearches)
End Sub

Public Sub Search(ByVal pstrPath As String, ByVal pstrExtension As String)
Dim objSearch   As Object
Dim objPB       As PropertyBag
Dim bytData()   As Byte
    Set objSearch = CreateObject("SearchThread.clsNewSearch")
    mobjGateway.StartLink objSearch.hWnd
    Searches.NewSearch Key, objSearch.hWnd
    Set objPB = New PropertyBag
    objPB.WriteProperty "PATH", pstrPath
    objPB.WriteProperty "EXTENSION", pstrExtension
    bytData = objPB.Contents
    Set objPB = Nothing
    mobjGateway.PostData bytData
    Set objSearch = Nothing
End Sub

Private Sub mobjGateway_DataArrived(Data As Variant, ByVal Synchronous As Boolean)
Dim objPB   As PropertyBag
On Error GoTo Errhandler
    Set objPB = New PropertyBag
    objPB.Contents = Data
    Searches.FileFound Key, objPB.ReadProperty("HWND"), objPB.ReadProperty("COUNT"), objPB.ReadProperty("FILE")
    Set objPB = Nothing
    Exit Sub
Errhandler:
    Beep
    'sometimes the byte array is slightly incorrect and causes an error.
    'still working on it.
End Sub

Private Sub mobjGateway_LinkTerminated()
    Destroy
End Sub

Private Sub Destroy()
    If Not (mobjGateway Is Nothing) Then
        Searches.RemoveSearch Key
        Set mobjGateway = Nothing
    End If
End Sub
