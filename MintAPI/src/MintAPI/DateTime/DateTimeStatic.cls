VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DateTimeStatic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'@PROJECT_LICENSE

''<class name="DateTimeStatic" static tags="">
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
Const CLASSID As String = "DateTimeStatic" '#DO NOT REMOVE IT
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
Public Function Create(ByVal Year As Long, ByVal Month As Long, ByVal Day As Long, _
                       ByVal Hour As Long, ByVal Minute As Long, ByVal Second As Long, ByVal Millisecond As Long) As DateTime
    Set Create = New DateTime
    Call Create.Constructor3(Year, Month, Day, Hour, Minute, Second, Millisecond)
End Function
Public Function CreateDate(ByVal Year As Long, ByVal Month As Long, ByVal Day As Long) As DateTime
    Set CreateDate = New DateTime
    Call CreateDate.Constructor3(Year, Month, Day, 0, 0, 0, 0)
End Function
Public Function CreateTime(ByVal Hour As Long, ByVal Minute As Long, ByVal Second As Long, ByVal Millisecond As Long) As DateTime
    Set CreateTime = New DateTime
    Call CreateTime.Constructor3(0, 0, 0, Hour, Minute, Second, Millisecond)
End Function

Public Function FromVBDate(ByVal Value As Date) As DateTime
    Set FromVBDate = New DateTime
    Call FromVBDate.Constructor2(Value)
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

Friend Property Get SystemTime() As API_SYSTEMTIME
    Call API_GetSystemTime(SystemTime)
End Property

Public Property Get Today() As DateTime
    Dim T As API_SYSTEMTIME
    T = SystemTime
    Set Today = New DateTime
    Call Today.Constructor3(T.wYear, T.wMonth, T.wDay, 0, 0, 0, 0)
End Property
Public Property Get Now() As DateTime
    Set Now = New DateTime
    Call Now.Constructor0
End Property
Public Property Get Yesterday() As DateTime
    Set Yesterday = Today.AddDays(-1)
End Property
Public Property Get Tommorow() As DateTime
    Set Tommorow = Today.AddDays(1)
End Property


Public Function Parse(ByVal Value As String) As DateTime
    
End Function
Public Function TryParse(ByVal Value As String, ByRef outRetVal As DateTime) As Boolean
    
End Function

'*********************************************
'*********************************************
'
'</section>
'---------------------------------------------
'---------------------------------------------
