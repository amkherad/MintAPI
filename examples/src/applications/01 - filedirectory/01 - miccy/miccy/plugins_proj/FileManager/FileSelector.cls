VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FileSelector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Implements IAction
Dim mp As masterProps

Private Sub Class_Initialize()
    Set mp = New masterProps
End Sub

Private Sub IAction_proccessAction(ByVal provider As miccyUltimateTools.provider)
    mp.SubFolders
End Sub
Private Function IAction_provideSettings() As ISettings
    Set IAction_provideSettings = mp
End Function
