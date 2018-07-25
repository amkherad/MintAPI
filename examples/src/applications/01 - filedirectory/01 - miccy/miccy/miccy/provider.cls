VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "provider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"Instance"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public Sub registerAction(ByVal IAction As IAction, actionUniqueID As String, actionName As String, actionDescrioption As String, Source As String, actionGroup As String)
    Call registerAction(IAction, actionUniqueID, actionName, actionDescrioption, Source, actionGroup)
End Sub
Public Sub registerFilter(ByVal IFilter As IFilter, filterUniqueID As String, filterName As String, filterDescription As String, Source As String, filterGroup As String)
    Call registerFilter(IFilter, filterUniqueID, filterName, filterDescription, Source, filterGroup)
End Sub

Public Function Process() As Process
    Set Process = lastProc
End Function
Public Function ActionBuffer() As ActionBuffer
    Set ActionBuffer = actionBuf
End Function
