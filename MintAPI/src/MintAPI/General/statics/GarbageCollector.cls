VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GarbageCollector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'@PROJECT_LICENSE

''<class name="GarbageCollector" static tags="">
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
Const CLASSID As String = "GarbageCollector" '#DO NOT REMOVE IT
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
Public Enum GarbageCollectorStates
    gcCollecting
    gcActive
    gcWaiting
    gcDead
End Enum
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

'Public Sub Collect(Optional TreatAsIdleTime As Boolean = False)
'    Dim Args As GarbageCollectorArguments
'    Args.Break = False
'    Args.Force = False
'    Args.IsEndOfTheApplication = False
'    Args.IsIdleTime = TreatAsIdleTime
'    Args.TicksPerModule = -1
'    Call mint_garbagec_BeginCollector(Args)
'End Sub

Public Sub Collect()
    
End Sub

Public Sub AddMemoryPressure(ByVal BytesAllocated As Long)
    
End Sub
Public Sub RemoveMemoryPressure(ByVal BytesAllocated As Long)
    
End Sub

Public Sub KeepAlive(ByVal Obj As Object)
    Dim IsIObject As Boolean
    IsIObject = TypeOf Obj Is IObject
    Call KeepObjAlive(Obj)
End Sub
Private Sub KeepObjAlive(ByVal Obj As Object)
    'Do nothing...
End Sub

Public Sub AddToFamilyChain(ByVal FamilyHolder As IObject, _
                            ByVal ObjToKeepAlive As Object)
    Dim AliveChain As List
    Set AliveChain = FamilyHolder.MetaObject.AliveChain
    If Not AliveChain.Contains(ObjToKeepAlive) Then _
        Call AliveChain.Append(ObjToKeepAlive)
End Sub
Public Sub RemoveFromFamilyChain(ByVal FamilyHolder As IObject, _
                            ByVal ObjToKeepAlive As Object)
    Dim AliveChain As List
    Set AliveChain = FamilyHolder.MetaObject.AliveChain
    If AliveChain.Contains(ObjToKeepAlive) Then _
        Call AliveChain.Remove(ObjToKeepAlive)
End Sub

'*********************************************
'*********************************************
'
'</section>
'---------------------------------------------
'---------------------------------------------
