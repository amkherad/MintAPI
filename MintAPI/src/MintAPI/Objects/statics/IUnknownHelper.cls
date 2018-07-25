VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IUnknownHelper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'@PROJECT_LICENSE

''<class name="" static tags="">
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
Const CLASSID As String = "MintAPIStaticClass" '#DO NOT REMOVE IT
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
Dim DesignMode_Refrences As List
'
'</section>
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section Creators>
'
Public Function IsIUnknown(ByRef Target As Variant) As Boolean
    IsIUnknown = IsObject(Target)
End Function
Public Function AsIUnknown(ByVal Obj As Object) As IUnknown
    Set AsIUnknown = Obj
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

Public Function AddRefPtr(ByVal pUnk As Long) As Long 'IUnknown::AddRef()
    If pUnk = vbNullPtr Then throw Exps.ArgumentNullException("pUnk")
    
    Dim TIVBUnknown As IVBUnknown
    Call memcpy(TIVBUnknown, pUnk, VLEN_OBJECT)
    AddRefPtr = TIVBUnknown.AddRef
    Call memzero(TIVBUnknown, VLEN_OBJECT)
End Function
Public Function ReleasePtr(ByVal pUnk As Long) As Long 'IUnknown::Release()
    If pUnk = vbNullPtr Then throw Exps.ArgumentNullException("pUnk")
    
    Dim TIVBUnknown As IVBUnknown
    Call memcpy(TIVBUnknown, pUnk, VLEN_OBJECT)
    ReleasePtr = TIVBUnknown.Release
    Call memzero(TIVBUnknown, VLEN_OBJECT)
End Function
Public Function QueryInterfacePtr(ByVal pUnk As Long, ByVal Guid As Guid, ByRef RetPtr As Long, ByRef RetObj As Object, Optional ByVal AutoAddRef As Boolean = False) As Long 'IUnknown::QueryInterface()
    If pUnk = vbNullPtr Then throw Exps.ArgumentNullException("pUnk")
    If Guid Is Nothing Then throw Exps.ArgumentNullException("Guid")
    Dim TIVBUnknown As IVBUnknown, RetRef As Long
    Call memcpy(TIVBUnknown, pUnk, VLEN_OBJECT)
    QueryInterfacePtr = TIVBUnknown.QueryInterface(Guid.ToStdGuid, RetRef)
    If QueryInterfacePtr <> S_OK Then Exit Function
    If AutoAddRef Then Call TIVBUnknown.AddRef
    Call memzero(TIVBUnknown, VLEN_OBJECT)
    Set RetObj = Nothing
    RetPtr = RetRef
    Call memcpy(RetObj, RetRef, VLEN_PTR)
End Function
Public Sub QueryInterfacePtr1(ByVal pUnk As Long, ByVal Guid As Guid, ByRef RetPtr As Long, ByRef RetObj As Object, Optional ByVal AutoAddRef As Boolean = True) 'IUnknown::QueryInterface()
    Dim Result As Long
    Result = QueryInterfacePtr(pUnk, Guid, RetPtr, RetObj, AutoAddRef)
    If Result <> S_OK Then throw Exps.InvalidCastException
End Sub

'==============================================


Public Function AddRef(ByVal UnkObject As Object) As Long  'IUnknown::AddRef()
    If UnkObject Is Nothing Then throw Exps.ArgumentNullException("UnkObject")
    If IsDesignMode Then
        If DesignMode_Refrences Is Nothing Then Set DesignMode_Refrences = New List
        Call DesignMode_Refrences.Append(UnkObject)
    Else
        AddRef = AddRefPtr(ObjPtr(UnkObject))
    End If
End Function
Public Function Release(ByVal UnkObject As Object) As Long 'IUnknown::Release()
    If UnkObject Is Nothing Then throw Exps.ArgumentNullException("UnkObject")
    If IsDesignMode Then
        If Not DesignMode_Refrences Is Nothing Then
            Call DesignMode_Refrences.Remove(UnkObject)
        End If
    Else
        Release = ReleasePtr(ObjPtr(UnkObject))
    End If
End Function

Public Function QueryInterface(ByVal UnkObject As Object, ByVal Guid As Guid, ByRef RetPtr As Long, ByRef RetObj As Object, Optional ByVal AutoAddRef As Boolean = True) As Long 'IUnknown::QueryInterface()
    If UnkObject Is Nothing Then throw Exps.ArgumentNullException("UnkObject")
    QueryInterface = QueryInterfacePtr(ObjPtr(UnkObject), Guid, RetPtr, RetObj, AutoAddRef)
End Function
Public Sub QueryInterface1(ByVal UnkObject As Object, ByVal Guid As Guid, ByRef RetPtr As Long, ByRef RetObj As Object, Optional ByVal AutoAddRef As Boolean = True) 'IUnknown::QueryInterface()
    If UnkObject Is Nothing Then throw Exps.ArgumentNullException("UnkObject")
    Dim Result As Long
    Result = QueryInterfacePtr(ObjPtr(UnkObject), Guid, RetPtr, RetObj, AutoAddRef)
    If Result <> S_OK Then throw Exps.InvalidCastException
End Sub


'*********************************************
'*********************************************
'
'</section>
'---------------------------------------------
'---------------------------------------------
