VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MintAPI2ndLayerAPI"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function LanguageEditor(Optional Path, Optional Parent) As LanguageEditor
    Dim le As New LanguageEditor
    Call le.Initialize(Path, Parent)
    Set LanguageEditor = le
End Function

Public Function ConfigurationEditor(Optional Path, Optional Parent) As ConfigurationEditor
    Dim ce As New ConfigurationEditor
    Call ce.Initialize(Path, Parent)
    Set ConfigurationEditor = ce
End Function

Public Function SourceCodeManager(FileFilters As GlobalFilters, Optional Path) As SourceCodeManager
    Dim scm As New SourceCodeManager
    Call scm.Initialize
    Set SourceCodeManager = scm
End Function

Public Sub EnableVisualStyles()
    Call Application.EnableVisualStyles
End Sub

Public Function IsInModalState() As Boolean
    IsInModalState = modMain.ModalState
End Function
