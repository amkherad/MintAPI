VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSearches"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Event SearchRemoved(ByVal pstrKey As String)
Public Event NewSearch(ByVal pstrKey As String, ByVal plnghWnd As Long)
Public Event FileFound(ByVal pstrKey As String, ByVal plnghWnd As Long, ByVal plngCount As Long, ByVal pstrCurrentFile As String)

Private mcolItems   As Collection

Private Sub Class_Initialize()
    Set mcolItems = New Collection
End Sub

Private Sub Class_Terminate()
    Set mcolItems = Nothing
End Sub

Public Property Get Item(ByVal Index As Variant) As clsSearch
    Set Item = mcolItems.Item(Index)
End Property

Public Property Get Count() As Long
    Count = mcolItems.Count
End Property

Public Sub CreateNew(ByVal pstrPath As String, ByVal pstrExtension As String)
Dim objSearch   As clsSearch
    Set objSearch = New clsSearch
    mcolItems.Add objSearch, objSearch.Key
    objSearch.Initialise Me
    objSearch.Search pstrPath, pstrExtension
    Set objSearch = Nothing
End Sub

Friend Sub NewSearch(ByVal pstrKey As String, ByVal plnghWnd As Long)
    RaiseEvent NewSearch(pstrKey, plnghWnd)
End Sub

Friend Sub RemoveSearch(ByVal pstrKey As String)
    RaiseEvent SearchRemoved(pstrKey)
    mcolItems.Remove pstrKey
End Sub

Friend Sub FileFound(ByVal pstrKey As String, ByVal plnghWnd As Long, ByVal plngCount As Long, ByVal pstrCurrentFile As String)
    RaiseEvent FileFound(pstrKey, plnghWnd, plngCount, pstrCurrentFile)
    'DoEvents
End Sub
