VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MMM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'@PROJECT_LICENSE

''<class name="MMM" static tags="">
''  <summary>
''    Provides functionality of MintAPI Memory Manager.
''  </summary>
''  <sample></sample>
''</class>

Option Base 0
Option Explicit

'=============================================
'=============================================
'=============================================
'<section Meta Data>
'
Const CLASSID As String = "MMM" '#DO NOT REMOVE IT
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
Private p_MMMRecords As New List
'
'</section>
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section Creators>
'
Public Function Alloc(ByVal Length As Long) As Long
    
End Function
Public Sub Free(ByVal Ptr As Long)
    Dim MMMR As MMMRecord ', Target As MMMRecord
    For Each MMMR In p_MMMRecords
        If MMMR.Ptr = Ptr Then
            'Set Target = MMMR
            Exit Sub
        End If
    Next
End Sub

Public Function AllocBlock(ByVal Length As Long) As IMemoryBlock
    Dim MMMB As New MMMBlock, Ptr As Long
    Ptr = Alloc(Length)
    If Ptr = vbNullPtr Then throw Exps.OutOfMemoryException
    Call MMMB.Constructor0(Ptr, Length, mapExecuteRead)
    Set AllocBlock = MMMB
End Function
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

'All your methods here...

'*********************************************
'*********************************************
'
'</section>
'---------------------------------------------
'---------------------------------------------

