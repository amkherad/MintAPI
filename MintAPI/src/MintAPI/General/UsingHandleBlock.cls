VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UsingHandleBlock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@PROJECT_LICENSE

''<class name="UsingHandleBlock" static tags="">
''  <summary>
''    Provides a mechanism to code like c# when working on IHandle derived classes.
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
Const CLASSID As String = "UsingHandleBlock" '#DO NOT REMOVE IT
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
'Public Function Create(ByVal Args As Long) As Object
'
'End Function
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

Private p_Obj As IHandle
Private p_Dispose As Boolean

Friend Sub Constructor0(ByVal Disposable As IHandle)
    If Not TypeOf Disposable Is IDisposable Then throw Exps.InvalidArgumentException("Object must implement IDisposable")
    Set p_Obj = Disposable
    p_Dispose = True
End Sub

Public Property Get Handle() As Long
    If p_Obj Is Nothing Then throw Exps.ObjectDisposedException
    Handle = p_Obj.Handle
End Property

Public Property Get Object() As IHandle
    If p_Obj Is Nothing Then throw Exps.ObjectDisposedException
    Set Object = p_Obj
End Property
'Public Property Let Object(ByVal Obj As IDisposable)
'    Set p_Obj = Obj
'End Property
'Public Property Set Object(ByVal Obj As IDisposable)
'    Set p_Obj = Obj
'End Property


Public Sub KeepAlive()
    Call GC.KeepAlive(Me)
End Sub

Public Sub Break()
    p_Dispose = False
End Sub

Private Sub Class_Terminate()
    If p_Dispose Then
        If Not p_Obj Is Nothing Then
            Dim Disp As IDisposable
            Set Disp = p_Obj
            Call Disp.Dispose
        End If
    End If
End Sub

'*********************************************
'*********************************************
'
'</section>
'---------------------------------------------
'---------------------------------------------


