VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Information"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"Application"
Attribute VB_Ext_KEY = "Member1" ,"Environment"
Attribute VB_Ext_KEY = "Member2" ,"Thread"
Attribute VB_Ext_KEY = "Member3" ,"Process"
'@PROJECT_LICENSE

''<class name="Information" static tags="">
''  <summary>
''    Provides some information about objects, current application and current environment.
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
Const CLASSID As String = "Information" '#DO NOT REMOVE IT
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
'Dim myAPI As API
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

Public Function IsRecord(ByRef Expression As Variant) As Boolean
    If IsMissing(Expression) Then Exit Function
    IsRecord = (VarType(Expression) = vbUserDefinedType)
End Function
Public Function IsByRef(Var) As Boolean
    IsByRef = ((VariantType(Var) And VT_BYREF) = VT_BYREF)
End Function
Public Function IsInteger(ByRef Arg As Variant) As Boolean
    Dim VT As VbVarType
    VT = VarType(Arg)
    Select Case VT
        Case vbByte, vbInteger, vbLong, vbDecimal ', vbDouble, vbSingle
            IsInteger = True
        Case Else
            IsInteger = False
    End Select
End Function
Public Function IsNumber(ByRef Arg As Variant) As Boolean
    Dim VT As VbVarType
    VT = VarType(Arg)
    Select Case VT
        Case vbByte, vbInteger, vbLong, vbDecimal, vbDouble, vbSingle
            IsNumber = True
        Case Else
            If VBA.IsNumeric(Arg) Then
                IsNumber = True
            Else
                IsNumber = False
            End If
    End Select
End Function


Public Function GetArgumentLength(ByVal VT As VariantTypes) As Long
    If (VT And VT_BYREF) = VT_BYREF Then _
        VT = VT Xor VT_BYREF
    Select Case VT
        Case VT_I4, VT_R4, VT_BSTR, VT_LPSTR, VT_LPWSTR, VT_DISPATCH, VT_ARRAY, VT_Ptr, VT_INT_PTR, VT_INT, VT_HRESULT, VT_BSTR_BLOB
            GetArgumentLength = VLEN_PTR
        Case VT_I8, VT_R8, VT_DECIMAL, VT_DATE
            GetArgumentLength = VLEN_DOUBLE
        Case VT_I2, VT_BOOL, VT_UI2
            GetArgumentLength = VLEN_BOOLEAN
        Case VT_VARIANT
            GetArgumentLength = VLEN_VARIANT
        Case VT_I1, VT_UI1: GetArgumentLength = VLEN_BYTE
        Case VT_UI4, VT_UINT, VT_UINT_PTR: GetArgumentLength = VLEN_LONG
        Case VT_UI8: GetArgumentLength = VLEN_DOUBLE
        Case VLEN_VARIANT: GetArgumentLength = VLEN_VARIANT
        Case VT_DECIMAL: GetArgumentLength = VLEN_DECIMAL
        Case Else
            If (VT And VT_ARRAY) = VT_ARRAY Then
                GetArgumentLength = VLEN_PTR
            Else
                GetArgumentLength = LenB(VT)
            End If
    End Select
End Function

'*********************************************
'*********************************************
'
'</section>
'---------------------------------------------
'---------------------------------------------
