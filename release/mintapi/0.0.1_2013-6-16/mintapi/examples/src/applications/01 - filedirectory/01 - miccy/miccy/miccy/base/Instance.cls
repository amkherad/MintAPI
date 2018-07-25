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
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Const CLASSID As String = "Instance"

Public Sub appendWindow(ByVal Iwindow As Iwindow): throw NotImplementedException: End Sub
Public Sub removeWindow(ByVal Iwindow As Iwindow): throw NotImplementedException: End Sub
Public Sub closeWindow(ByVal Iwindow As Iwindow): throw NotImplementedException: End Sub
Public Sub showWindow(ByVal Iwindow As Iwindow): throw NotImplementedException: End Sub
Public Sub temperaryshowWindow(ByVal Iwindow As Iwindow): throw NotImplementedException: End Sub
Public Function canShowWindow() As Boolean: throw NotImplementedException: End Sub
Public Sub appendMenu(Menu As String, Caption As String, Command As String): throw NotImplementedException: End Sub
Public Sub removeMenu(Menu As String): throw NotImplementedException: End Sub
Public Sub searchPlugin(PluginName As String): throw NotImplementedException: End Sub
Public Sub installPlugin(PluginPath As String): throw NotImplementedException: End Sub
Public Sub uninstallPlugin(PluginUniqueID As String): throw NotImplementedException: End Sub

Public Sub registerAction(ByVal IAction As IAction, actionUniqueID As String, actionName As String, actionDescrioption As String, Source As String, actionGroup As String): throw NotImplementedException: End Sub
Public Sub unregisterAction(actionUniqueID As String): throw NotImplementedException: End Sub

Public Sub registerFilter(ByVal IFilter As IFilter, filterUniqueID As String, filterName As String, filterDescription As String, Source As String, filterGroup As String): throw NotImplementedException: End Sub
Public Sub unregisterFilter(filterUniqueID As String): throw NotImplementedException: End Sub

Public Property Get countActions() As Long: throw NotImplementedException: End Sub
Public Property Get countFilters() As Long: throw NotImplementedException: End Sub
Public Property Get countPlugins() As Long: throw NotImplementedException: End Sub

Public Function proccess() As IProccess: throw NotImplementedException: End Function
Public Sub cerr(Message As String): throw NotImplementedException: End Sub
Public Sub clog(Message As String): throw NotImplementedException: End Sub

Public Function AskQuestion(Title As String, Description As String, flags As Long) As VbMsgBoxResult: throw NotImplementedException: End Function
