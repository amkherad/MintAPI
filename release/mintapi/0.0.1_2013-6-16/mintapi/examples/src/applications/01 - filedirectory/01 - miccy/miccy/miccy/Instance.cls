VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Instance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Const CLASSID As String = "Instance"

Public Sub cerr(ByVal Message As String)
    Call modMain.cerr(Message)
End Sub
Public Sub clog(ByVal Message As String)
    Call modMain.clog(Message)
End Sub
Public Sub cdebug(ByVal Message As String)
    Call modMain.cdebug(Message)
End Sub
Public Sub cfatal(ByVal Message As String)
    Call modMain.cfatal(Message)
End Sub
Public Sub cthrow(ByVal Message As String)
    Call modMain.cthrow(Message)
End Sub


Public Function AskQuestion(title As String, Description As String, flags As Long) As VbMsgBoxResult
    AskQuestion = MsgBox(Description, flags, title)
End Function

Public Function canShowWindow() As Boolean
    
End Function

Public Property Get countActions() As Long
    
End Property

Public Property Get countFilters() As Long
    
End Property

Public Property Get countPlugins() As Long
    
End Property

Public Sub installPlugin(PluginPath As String)
    
End Sub

Public Function Process() As Process
    Set Process = lastProc
End Function

Public Sub registerAction(ByVal IAction As IAction, actionUniqueID As String, actionName As String, actionDescrioption As String, Source As String, actionGroup As String)
    
End Sub

Public Sub registerFilter(ByVal IFilter As IFilter, filterUniqueID As String, filterName As String, filterDescription As String, Source As String, filterGroup As String)
    
End Sub

Public Sub appendWindow(ByVal Iwindow As Iwindow)
    Call Iwindows.Add(Iwindow)
End Sub
Public Sub removeWindow(ByVal Iwindow As Iwindow)
    Dim i As Long
    For i = 0 To Iwindows.Count - 1
        Call Iwindows.Remove(i)
    Next
End Sub
Public Sub appendMenu(menu As String, Caption As String, pluginCallBack As IPlugin, Command As String)
    
End Sub
Public Sub removeMenu(menu As String)
    
End Sub

Public Sub searchPlugin(PluginName As String)
    
End Sub
Public Sub ShowWindow(ByVal Iwindow As Iwindow)
    Dim modal As Boolean
    
End Sub
Public Sub CloseWindow(ByVal Iwindow As Iwindow)
    
End Sub
Public Sub temperaryshowWindow(ByVal Iwindow As Iwindow)
    
End Sub
Public Sub uninstallPlugin(PluginUniqueID As String)
    
End Sub
Public Sub unregisterAction(actionUniqueID As String)
    
End Sub
Public Sub unregisterFilter(filterUniqueID As String)
    
End Sub

