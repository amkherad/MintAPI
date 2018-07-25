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

Public Sub RegisterInstance(ByVal Instance As Instance, reserved) 'reserved must be Instance ptr -1
    If ObjPtr(Instance) <> (reserved - 1) Then Exit Sub
    Set glbInstance = Instance
End Sub
Public Sub registerAction(ByVal IAction As IAction, actionUniqueID As String, actionName As String, actionDescrioption As String, Source As String, actionGroup As String)
    Call glbInstance.registerAction(IAction, actionUniqueID, actionName, actiondescription, Source, actionGroup)
End Sub
Public Sub registerFilter(ByVal IFilter As IFilter, filterUniqueID As String, filterName As String, filterDescription As String, Source As String, filterGroup As String)
    Call glbInstance.registerFilter(IFilter, filterUniqueID, filterName, filterDescription, Source, filterGroup)
End Sub
Public Function isInstanceRegistered() As Boolean
    isInstanceRegistered = Not (glbInstance Is Nothing)
End Function
Public Function Instance() As Instance
    Set Instance = glbInstance
End Function
