VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ThreadStackStatic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'@PROJECT_LICENSE

''<class name="ThreadStackStatic" static tags="onlystdcall">
''  <summary>
''    Some class description.
''  </summary>
''  <sample></sample>
''</class>

'##################@@@@@@@@@@@@@@@@@@ Function replacement - becarefull!

Option Base 0
Option Explicit

'=============================================
'=============================================
'=============================================
'<section Meta Data>
'
Const CLASSID As String = "ThreadStackStatic" '#DO NOT REMOVE IT
Const CLASSALIAS As String = APPLICATIONDOMAIN & "." & CLASSID '#DO NOT REMOVE IT
Const CLASSVERSIONSTRING As String = "0.1" '#DO NOT REMOVE IT
Const CLASSALIASSTRING As String = CLASSALIAS & "/" & CLASSVERSIONSTRING '#DO NOT REMOVE IT

Const pFUNC_ORDER_StackModifiers = FUNC_ORDER_SKIP_IDISPATCH_B1 + 1

Const pFUNC_ORDER_PushLong As Long = pFUNC_ORDER_StackModifiers + 1
Const pFUNC_ORDER_PushFloat As Long = pFUNC_ORDER_StackModifiers + 2
Const pFUNC_ORDER_PushDouble As Long = pFUNC_ORDER_StackModifiers + 3
Const pFUNC_ORDER_PushCurrency As Long = pFUNC_ORDER_StackModifiers + 4
Const pFUNC_ORDER_PushString As Long = pFUNC_ORDER_StackModifiers + 5
Const pFUNC_ORDER_PushObject As Long = pFUNC_ORDER_StackModifiers + 6

Const pFUNC_ORDER_PopLong As Long = pFUNC_ORDER_StackModifiers + 7
Const pFUNC_ORDER_PopFloat As Long = pFUNC_ORDER_StackModifiers + 8
Const pFUNC_ORDER_PopDouble As Long = pFUNC_ORDER_StackModifiers + 9
Const pFUNC_ORDER_PopCurrency As Long = pFUNC_ORDER_StackModifiers + 10
Const pFUNC_ORDER_PopObject As Long = pFUNC_ORDER_StackModifiers + 11

Const pFUNC_ORDER_Reserve As Long = pFUNC_ORDER_PopObject + 1
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
Public Property Get Current() As ThreadStack
    'Thread.Current.StackBounderies.
End Property
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

Public Sub PushLong(ByVal lngData As Long): End Sub
Public Sub PushFloat(ByVal fltData As Single): End Sub
Public Sub PushDouble(ByVal dblData As Double): End Sub
Public Sub PushCurrency(ByVal curData As Currency): End Sub
Public Sub PushString(ByVal strData As String): End Sub
Public Sub PushObject(ByVal objData As Object): End Sub

Public Sub PopLong(ByRef lngData As Long): End Sub
Public Sub PopFloat(ByRef fltData As Single): End Sub
Public Sub PopDouble(ByRef dblData As Double): End Sub
Public Sub PopCurrency(ByRef curData As Currency): End Sub
Public Sub PopString(ByRef strData As String): End Sub
Public Sub PopObject(ByRef objData As Object): End Sub

Public Sub Reserve(ByVal Length As Long): End Sub

Public Property Get StackPointer() As Long
    StackPointer = mHelper.ReadCalleeEBP - 12 'EBP - RetPtr(4) - ObjRef(4) - RetVal(4)
End Property
Public Property Get BasePointer() As Long
    BasePointer = mHelper.ReadCallerEBP
End Property

'*********************************************
'*********************************************
'
'</section>
'---------------------------------------------
'---------------------------------------------

Private Sub Class_Initialize()
    Dim VTbl As Long, MRetPtr As Long
    Call memcpy(VTbl, ByVal ObjPtr(Me), VLEN_PTR)
    MRetPtr = GetMintAssembliesMethodGlobalInfo(FUNC_ORDER_RETURN).Reference
    Call memcpy(ByVal (VTbl + (pFUNC_ORDER_PushLong * VLEN_PTR)), MRetPtr, VLEN_PTR)
    Call memcpy(ByVal (VTbl + (pFUNC_ORDER_PushFloat * VLEN_PTR)), MRetPtr, VLEN_PTR)
    Call memcpy(ByVal (VTbl + (pFUNC_ORDER_PushDouble * VLEN_PTR)), MRetPtr, VLEN_PTR)
    Call memcpy(ByVal (VTbl + (pFUNC_ORDER_PushCurrency * VLEN_PTR)), MRetPtr, VLEN_PTR)
    Call memcpy(ByVal (VTbl + (pFUNC_ORDER_PushString * VLEN_PTR)), MRetPtr, VLEN_PTR)
    Call memcpy(ByVal (VTbl + (pFUNC_ORDER_PushObject * VLEN_PTR)), MRetPtr, VLEN_PTR)
    
    Call memcpy(ByVal (VTbl + (pFUNC_ORDER_Reserve * VLEN_PTR)), GetMintAssembliesMethodGlobalInfo(FUNC_ORDER_RESERVE).Reference, VLEN_PTR)
End Sub
