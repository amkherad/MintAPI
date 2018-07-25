VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "default"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Implements IPageCodeBehind
Const CLASSID As String = "default"

Dim r As Response
Dim rq As Request

Private Sub IPageCodeBehind_Construct()
    
End Sub
Private Sub IPageCodeBehind_Destruct()
    
End Sub

Private Sub IPageCodeBehind_OnOnUnLoad(ByVal e As zaX.EventArgs)
    
End Sub

Private Sub IPageCodeBehind_Request(ByVal Response As zaX.Response, ByVal Request As zaX.Request)
    Set r = Response
    Set rq = Request
End Sub
Private Sub IPageCodeBehind_OnLoad(ByVal e As zaX.EventArgs)
    out r, "<html>"
        out r, "<head>"
            out r, "<title>", "zaX Server Sample Website", "</title>"
        out r, "</head>"
        out r, "<body>"
            out r, "<h1>", "this is sample text.", "</h1>"
            out r, "<br />"
            out r, "<h4>", "this is another sample text.", "</h4>"
            out r, parse("new widgets.button()", Cprops("text:Please Click Me!"), Cevent("onclick:zaXS.button01_click"))
            outp r, "new widgets.button()", Cprops("text:Please Click Me!"), Cevent("onclick:zaXS.button01_click")
        out r, "</body>"
    out r, "</html>"
End Sub
Private Sub IPageCodeBehind_OnFinishLoad(ByVal e As zaX.EventArgs)
    
End Sub
Private Sub IPageCodeBehind_OnUnLoad(ByVal e As zaX.EventArgs)
    
End Sub

