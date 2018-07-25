VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ArgumentListStatic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'@PROJECT_LICENSE

''<class name="ArgumentListStatic" static tags="onlystdcall">
''  <summary>
''    Some class description.
''  </summary>
''  <remarks>Use these functions only in STDCall methods.</remarks>
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
Const CLASSID As String = "ArgumentListStatic" '#DO NOT REMOVE IT
Const CLASSALIAS As String = APPLICATIONDOMAIN & "." & CLASSID '#DO NOT REMOVE IT
Const CLASSVERSIONSTRING As String = "0.1" '#DO NOT REMOVE IT
Const CLASSALIASSTRING As String = CLASSALIAS & "/" & CLASSVERSIONSTRING '#DO NOT REMOVE IT

'Const FUNC_ORDER_FromParamArray As Long = FUNC_ORDER_SKIP_IDISPATCH_B1 + 2
'
'</section>
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section API/Type Declarations>
'

Public Enum ParameterSizes
    vsByVal = 0
    vsByte = 4
    vsBoolean = 4
    vsInteger = 4
    vsInt = 4
    vsLong = 4
    vsShort = 4
    vsLongLong = 8
    vsDouble = 8
    vsSingle = 4
    vsFloat = 4
    vsCurrency = 8
    vsDate = 8
    
    vsPointer = 4
    vsDecimal = vsPointer
    vsString = vsPointer
    vsObject = vsPointer
    vsDataObject = vsPointer
    vsVariant = vsPointer
    vsVariantReference = vsPointer
    vsByRef = vsPointer
    vsUDTReference = vsPointer
    vsArray = vsPointer
    vsParamArray = vsPointer
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
Public Function Create(ParamArray Args() As Variant) As ArgumentList
Attribute Create.VB_UserMemId = 0
    Set Create = New ArgumentList
    Dim pArgs() As Variant
    pArgs = Args
    Call Create.Constructor2(pArgs)
    Call Create.Localize
    Call memzeroarr(pArgs, VLEN_PTR)
End Function
''<summary>
''  Provides access to a ParamArray argument to pass it to a sub calling procedure.
''<br>
''</summary>
''<sample>
''  Public Function MyParamArrayMethod(ParamArray SomeArgs() As Variant)
''      Call myParamArrayPassedMethod(AL.FromParamArray(0).ToArray())
''  End Function
''</sample>
''<remarks>
''  It's very important to use this method and created object ONLY in
''  calling procedure and don't store it outside of procedure and it's
''  important to terminate the created object at the end of procedure.
''<br>
''  ex:<br>
''<code>
''  Public Function MyParamArrayMethod(ParamArray SomeArgs() As Variant)
''      'Then try to call a method with following prototype:
''      '  Function myParamArrayPassedMethod(ByRef Array As Variant)
''      Call myParamArrayPassedMethod(AL.FromParamArray(0).ToArray())
''      'or:
''      Dim myVar As Variant 'it can be array: myVar() or blob: myVar(3) 'read Vector's CopyTo() or ArgumentList CopyTo()/MoveTo()
''      Call AL.FromParamArray(0).MoveTo(myVar)
''      Call myParamArrayPassedMethod(myVar)
''      'Call AL.FromParamArray(0).MoveTo(myGLOBALVar) is wrong!
''  End Function
''</code>
''  It happens because FromParamArray() based on a mechanism that don't clone
''  the original keeper array of arguments and just point to that, so at the
''  end of procedure this array will reclaim from memory and ArgumentList will
''  point to a wrong address and any access to it may fail (cause access violation crash)
''</remarks>
Public Function FromParamArray(Optional ByVal Index_InBytes As ParameterSizes = 0) As ArgumentList
    'EBP  RET  OBJ  VAR1   ... VARn
    ' |xxxx|xxxx|xxxx|xxxx ...  |xxxx
    Dim PEBP As Long 'Parent EBP
    PEBP = (mHelper.ReadCallerEBP + &HC) + Index_InBytes

    Set FromParamArray = New ArgumentList
    Call FromParamArray.Constructor0(PEBP)
End Function
Public Function FromArray(TArray, Optional ByVal Clone As Boolean = True) As ArgumentList
    Set FromArray = New ArgumentList
    If Clone Then
        Call FromArray.Constructor3(TArray)
    Else
        Call FromArray.Constructor2(TArray)
    End If
End Function
Public Function FromDictionary(ByVal Dictionary As IDictionary) As ArgumentList
    
End Function
'
'</section>
'---------------------------------------------
'---------------------------------------------
