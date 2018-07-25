VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SignalStatic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'@PROJECT_LICENSE

''<class name="SignalStatic" static tags="">
''  <summary>
''    Some class description.
''  </summary>
''  <using>Some class alias.</using>
''  <idea>Me , Qt</idea>
''  <sample></sample>
''</class>

Option Base 0
Option Explicit

'=============================================
'=============================================
'=============================================
'<section Meta Data>
'
Const CLASSID As String = "SignalStatic" '#DO NOT REMOVE IT
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

Public Function Emit(ByVal SignalObject As IObject, ByRef Signal As Variant, ParamArray Args() As Variant) As Variant
    'Call API_VariantCopy(Emit, GetEmitterOf(SignalObject).Emit(Signal, AL.FromParamArray(8)))
End Function
Public Function EmitA(ByVal SignalObject As IObject, ByRef Signal As Variant, Args() As Variant) As Variant
    'Call API_VariantCopy(EmitA, GetEmitterOf(SignalObject).Emit(Signal, AL.FromArray(Args)))
End Function
Public Function EmitAL(ByVal SignalObject As IObject, ByRef Signal As Variant, ByVal Args As ArgumentList) As Variant
    'Call API_VariantCopy(EmitAL, GetEmitterOf(SignalObject).Emit(Signal, Args))
End Function

Public Sub Connect(ByRef Signal As Variant, ByVal SignalObject As IObject, ByVal Slot As Method)
    
End Sub
Public Sub Disconnect(ByRef Signal As Variant, ByVal SignalObject As IObject, ByVal Slot As Method)
    
End Sub

Public Property Get SignalsBlockState(ByVal SignalObject As IObject) As Boolean
    SignalsBlockState = GetEmitterOf(SignalObject).SignalsState
End Property
Public Property Let SignalsBlockState(ByVal SignalObject As IObject, ByVal Value As Boolean)
    GetEmitterOf(SignalObject).SignalsState = Value
End Property


Private Function GetEmitterOf(ByVal SigObj As IObject) As SignalEmitter
    If SigObj Is Nothing Then throw Exps.ArgumentNullException
    Dim MO As MetaObject
    Set MO = SigObj.MetaObject
    If MO Is Nothing Then throw Exps.InvalidOperationException("An Error Occured When Trying To Retrieve The MetaObject Of The Class.")
    Set GetEmitterOf = MO.SignalEmitter
    If GetEmitterOf Is Nothing Then throw Exps.InvalidOperationException("An Error Occured When Trying To Retrieve The SignalEmitter Of The Class.")
End Function


'*********************************************
'*********************************************
'
'</section>
'---------------------------------------------
'---------------------------------------------
