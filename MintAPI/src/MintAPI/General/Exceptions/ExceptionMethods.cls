VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ExceptionMethods"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'@PROJECT_LICENSE

''<class name="ExceptionMethods" static tags="">
''  <summary>
''    Provides general exception handling and exception based debugging methods.
''  </summary>
''  <using>Exception</using>
''  <idea>Me</idea>
''  <sample>
''      throw Exps.Exception("Some message.")
''  </sample>
''</class>

Option Base 0
Option Explicit

'=============================================
'=============================================
'=============================================
'<section Meta Data>
'
Const CLASSID As String = "ExceptionMethods" '#DO NOT REMOVE IT
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
'<section Class Public Body Members>
'
'*********************************************
'*********************************************

Public Sub throw(ByVal Exception As Exception, Optional ByVal AtMethod As String, Optional ByVal Details As String)
    If Exception Is Nothing Then throw Exps.ArgumentNullException("Exception")
    Dim Obj As Object, ObjPtr As Long
    ObjPtr = mHelper.CallerThis
    If ObjPtr <> vbNullPtr Then
        Call memcpy(Obj, ObjPtr, VLEN_PTR)
        Call IUnknown_AddRef(Obj)
    End If
    Call throw0(Exception, Obj, AtMethod, Details)
End Sub
Public Sub rethrow()
    Call mint_exceptions.rethrow
End Sub

Public Sub Ensure(ByVal Condition As Boolean, ByVal Exception As Exception)
    If Exception Is Nothing Then throw Exps.ArgumentNullException("Exception")
    If Not Condition Then
        Dim Obj As Object, ObjPtr As Long
        ObjPtr = mHelper.CallerThis
        If ObjPtr <> vbNullPtr Then
            Call memcpy(Obj, ObjPtr, VLEN_PTR)
            Call IUnknown_AddRef(Obj)
        End If
        Call mint_exceptions.throw(Exception)
    End If
End Sub
Public Sub Assert(ByVal Condition As Boolean, ByVal Exception As Exception)
    If Exception Is Nothing Then throw Exps.ArgumentNullException("Exception")
    If Not Condition Then
        Dim Obj As Object, ObjPtr As Long
        ObjPtr = mHelper.CallerThis
        If ObjPtr <> vbNullPtr Then
            Call memcpy(Obj, ObjPtr, VLEN_PTR)
            Call IUnknown_AddRef(Obj)
        End If
        Call mint_exceptions.throw(Exception)
    End If
End Sub
Public Sub AssertM(ByVal Condition As Boolean, ByVal Message As String)
    If Not Condition Then
        Dim Obj As Object, ObjPtr As Long
        ObjPtr = mHelper.CallerThis
        If ObjPtr <> vbNullPtr Then
            Call memcpy(Obj, ObjPtr, VLEN_PTR)
            Call IUnknown_AddRef(Obj)
        End If
        Call mint_exceptions.throw(Exps.Exception(Message))
    End If
End Sub

Public Function Catch(ByRef Exception As Exception) As Boolean
    If Not LastException.Catched Then
        Dim Exp As Exception
        Set Exp = LastException.Exception
        If Exp Is Nothing Then Exit Function
        Set LastException.Exception = Exp
        LastException.Catched = True
        Catch = True
    End If
End Function

'*********************************************
'*********************************************
'
'</section>
'---------------------------------------------
'---------------------------------------------
