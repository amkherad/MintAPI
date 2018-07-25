VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Method"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"Thread"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Base 0
Option Explicit
Const CLASSID As String = "Method"
Const CLASSALIAS As String = APPLICATIONDOMAIN & "." & CLASSID
Const CLASSVERSIONSTRING As String = "0.1"
Const CLASSALIASSTRING As String = CLASSALIAS & "/" & CLASSVERSIONSTRING

Private Declare Function API_IsBadCodePtr Lib "kernel32" Alias "IsBadCodePtr" (ByVal lpfn As Long) As Long

Implements IAliasable

Public Enum MethodTypes
    mtFunction
    mtMethod
End Enum

Dim nm As String
Dim targetFunction As Long
Dim lastConcurrentInvokeThread As Thread

Public Sub Initialize(ByVal Name As String, ByVal targetSiteFunction As Long)
    nm = Name
    targetFunction = targetSiteFunction
End Sub

Public Property Get Handle() As Long
    Handle = targetFunction
End Property
Public Property Get Name() As String
    Name = nm
End Property
Public Function Executable() As Boolean
    Executable = ((targetFunction <> 0) And (API_IsBadCodePtr(targetFunction) = False))
End Function

Private Function InvokeAsFunction(ByVal targetFunction As Long, Args() As Variant)
    Dim C As Long
    C = ArraySize(Args)
    Select Case C
        Case 0
            'InvokeAsFunction = API_CallFunction0(targetFunction)
        Case 1
            'InvokeAsFunction = API_CallFunction1(targetFunction, Args(0))
        Case 2
            'InvokeAsFunction = API_CallFunction2(targetFunction, Args(0), Args(1))
        Case 3
            'InvokeAsFunction = API_CallFunction3(targetFunction, Args(0), Args(1), Args(2))
        Case 4
            'InvokeAsFunction = API_CallFunction4(targetFunction, Args(0), Args(1), Args(2), Args(3))
        Case 5
            'InvokeAsFunction = API_CallFunction5(targetFunction, Args(0), Args(1), Args(2), Args(3), Args(4))
        Case 6
            'InvokeAsFunction = API_CallFunction6(targetFunction, Args(0), Args(1), Args(2), Args(3), Args(4), Args(5))
        Case 7
            'InvokeAsFunction = API_CallFunction7(targetFunction, Args(0), Args(1), Args(2), Args(3), Args(4), Args(5), Args(6))
        Case 8
            'InvokeAsFunction = API_CallFunction8(targetFunction, Args(0), Args(1), Args(2), Args(3), Args(4), Args(5), Args(6), Args(7))
        Case 9
            'InvokeAsFunction = API_CallFunction9(targetFunction, Args(0), Args(1), Args(2), Args(3), Args(4), Args(5), Args(6), Args(7), Args(8))
        Case 10
            'InvokeAsFunction = API_CallFunction10(targetFunction, Args(0), Args(1), Args(2), Args(3), Args(4), Args(5), Args(6), Args(7), Args(8), Args(9))
        Case Else
            throw InvalidCallException("Unable to invoke functions with more than 10 arguments.")
    End Select
End Function
Private Sub InvokeAsMethod(ByVal targetFunction As Long, Args() As Variant)
    Dim C As Long
    C = ArraySize(Args)
    Select Case C
        Case 0
            'Call API_CallMethod0(targetFunction)
        Case 1
            'Call API_CallMethod1(targetFunction, Args(0))
        Case 2
            'Call API_CallMethod2(targetFunction, Args(0), Args(1))
        Case 3
            'Call API_CallMethod3(targetFunction, Args(0), Args(1), Args(2))
        Case 4
            'Call API_CallMethod4(targetFunction, Args(0), Args(1), Args(2), Args(3))
        Case 5
            'Call API_CallMethod5(targetFunction, Args(0), Args(1), Args(2), Args(3), Args(4))
        Case 6
            'Call API_CallMethod6(targetFunction, Args(0), Args(1), Args(2), Args(3), Args(4), Args(5))
        Case 7
            'Call API_CallMethod7(targetFunction, Args(0), Args(1), Args(2), Args(3), Args(4), Args(5), Args(6))
        Case 8
            'Call API_CallMethod8(targetFunction, Args(0), Args(1), Args(2), Args(3), Args(4), Args(5), Args(6), Args(7))
        Case 9
            'Call API_CallMethod9(targetFunction, Args(0), Args(1), Args(2), Args(3), Args(4), Args(5), Args(6), Args(7), Args(8))
        Case 10
            'Call API_CallMethod10(targetFunction, Args(0), Args(1), Args(2), Args(3), Args(4), Args(5), Args(6), Args(7), Args(8), Args(9))
        Case Else
            throw InvalidCallException("Unable to invoke functions with more than 10 arguments.")
    End Select
End Sub
Public Function LocalInvoke(ByVal MethodType As MethodTypes, ParamArray Args() As Variant) As Long
    If Not Executable Then throw InvalidCallException("Target method is not executable.")
    Dim Arg() As Variant
    Dim i As Long, mx As Long
    On Error GoTo zeroLength
    mx = UBound(Args) + 1
zeroLength:
    For i = 0 To mx - 1
        If VarType(Args(i)) = VBObject Then
            Set Arg(i) = Args(i)
        Else
                Arg(i) = Args(i)
        End If
    Next
    If MethodType = mtFunction Then
        LocalInvoke = InvokeAsFunction(targetFunction, Arg)
    Else
        Call InvokeAsMethod(targetFunction, Arg)
    End If
End Function
Public Function Invoke(Optional Args)
    If Not Executable Then throw InvalidCallException("Target method is not executable.")
    Call Thread.Invoke(Args, Me)
End Function
Public Function ConcurrentInvoke(Optional Args) As Thread
    If Not Executable Then throw InvalidCallException("Target method is not executable.")
    Set lastConcurrentInvokeThread = Thread.ConcurrentInvoke(Args, Me)
End Function

Public Function Thread() As Thread
    If lastConcurrentInvokeThread Is Nothing Then
        Set Thread = mint_constructor.Thread(Me)
    Else
        Set Thread = lastConcurrentInvokeThread
    End If
End Function


Private Property Get IAliasable_Alias() As String
    IAliasable_Alias = CLASSALIASSTRING
End Property
