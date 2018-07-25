VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Interlocked"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'@PROJECT_LICENSE

''<class name="Interlocked" static tags="">
''  <summary>
''    The Interlocked class provides the mechanism for synchronizing access to a variable that is shared by multiple threads. The threads of different processes can use this mechanism if the variable is in shared memory.
''  </summary>
''  <using></using>
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
Const CLASSID As String = "Interlocked" '#DO NOT REMOVE IT
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

Public Function Increment(ByRef Value As Long) As Long
    Increment = API_InterlockedIncrement(Value)
End Function
Public Function Decrement(ByRef Value As Long) As Long
    Decrement = API_InterlockedDecrement(Value)
End Function
Public Function Exchange(ByRef Target As Long, ByVal Value As Long) As Long
    Exchange = API_InterlockedExchange(Target, Value)
End Function
Public Function ExchangeAdd(ByRef Target As Long, ByVal Value As Long) As Long
    ExchangeAdd = API_InterlockedExchangeAdd(Target, Value)
End Function

Public Sub CompareExchange(ByRef Destination As Long, ByRef Exchange As Long, ByRef Comperand As Long)
    Call API_InterlockedCompareExchange(Destination, Exchange, Comperand)
End Sub

'*********************************************
'*********************************************
'
'</section>
'---------------------------------------------
'---------------------------------------------
