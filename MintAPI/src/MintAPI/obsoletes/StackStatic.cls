VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StackStatic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'@PROJECT_LICENSE

''<class name="StackStatic" static tags="">
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
Const CLASSID As String = "StackStatic" '#DO NOT REMOVE IT
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
Public Function CreateStaticStack(ByVal Items As Long) As Stack
    
End Function
Public Function Create() As Stack
    
End Function
Public Function CreateReadOnlyStack(ByVal Stack As Stack, Optional ByVal ModifySource As Boolean = False) As Stack
    
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

Public Property Get AddressOfStack() As Long
    
End Property

''<summary>
''  Pushes a data to thread stack.
''</summary>
''<remarks>
''  The caller is responsible to Pop the data from stack,
''  otherwise the calling method can't return and this cause a system crash.
''</remarks>
Public Function Push(ByRef Bytes() As Byte, Optional ByVal Thread As Thread = Nothing) As Variant
    
End Function
''<summary>
''  Pop a data from thread stack, the data will be removed from stack.
''<summary>
''<remarks>
''  The caller is responsible to calling method's return pointer not to be removed,
''  otherwise the calling method can't return and this cause a system crash.
''<remarks>
Public Function Pop(ByRef Bytes() As Byte, ByVal Length As Long, Optional ByVal Thread As Thread = Nothing) As Variant
    
End Function
''<summary>
''  Peek a data from thread stack, the source data in stack will be untouched.
''<summary>
''<exceptions>
''  <exception name="">This will be thrown when ReadType length and Length have conflicts.</exception>
''</exceptions>
Public Function Peek(ByRef Bytes() As Byte, ByVal Length As Long, Optional ByVal Thread As Thread = Nothing) As Variant
    
End Function

'*********************************************
'*********************************************
'
'</section>
'---------------------------------------------
'---------------------------------------------
