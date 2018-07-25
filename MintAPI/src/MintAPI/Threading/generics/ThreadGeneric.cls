VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ThreadGeneric"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'@PROJECT_LICENSE

''<class name="ThreadGeneric" static tags="immutable">
''  <summary>
''    Some class description.
''  </summary>
''  <using>Some class alias.</using>
''  <idea>Me</idea>
''  <sample></sample>
''</class>

Option Base 0
Option Explicit

'=============================================
'=============================================
'=============================================
'<section Meta Data>
'
Const CLASSID As String = "ThreadGeneric" '#DO NOT REMOVE IT
Const CLASSALIAS As String = APPLICATIONDOMAIN & "." & CLASSID '#DO NOT REMOVE IT
Const CLASSVERSIONSTRING As String = "0.1" '#DO NOT REMOVE IT
Const CLASSALIASSTRING As String = CLASSALIAS & "/" & CLASSVERSIONSTRING '#DO NOT REMOVE IT

'
'</section>
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section API/Type Declarations>
'

'
'</section>
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section Creators>
'

'
'</section>
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section Class Public Body Members>
'
'*********************************************
'*********************************************

Public Function Await(ByVal AsyncResult As IAsyncResult) As Variant
    Call API_VariantCopy(Await, mint_threading.Await(AsyncResult))
End Function
Public Sub Async(): Call mint_threading.Async: End Sub

Public Sub LockObj(ByVal Obj As IObject)
    If Obj Is Nothing Then throw Exps.ArgumentNullException
    Call API_EnterCriticalSection(Obj.MetaObject.Synchronization.SyncHandle)
End Sub
Public Function TryLockObj(ByVal Obj As IObject) As Boolean
    If Obj Is Nothing Then throw Exps.ArgumentNullException
    If API_TryEnterCriticalSection(Obj.MetaObject.Synchronization.SyncHandle) <> NO_VALUE Then _
        Exit Function
    TryLockObj = True
End Function
Public Sub EndLockObj(ByVal Obj As IObject)
    If Obj Is Nothing Then throw Exps.ArgumentNullException
    Call API_LeaveCriticalSection(Obj.MetaObject.Synchronization.SyncHandle)
End Sub

'*********************************************
'*********************************************
'
'</section>
'---------------------------------------------
'---------------------------------------------
