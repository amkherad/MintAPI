VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "App"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Implements IWebApplication
    
Private Sub IWebApplication_Initialize()
    
End Sub
Private Sub IWebApplication_Terminate()
    
End Sub

Private Sub IWebApplication_Request(ByVal r As zaX.Response, ByVal req As zaX.Request)
    r.WriteParserObject parse("")
End Sub
