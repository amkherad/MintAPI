VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CallingConvension"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'@PROJECT_LICENSE

''<class name="CallingConvension" static tags="">
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
Const CLASSID As String = "CallingConvension" '#DO NOT REMOVE IT
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
'Public Function Create() As MethodCallingConvension
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

Public Property Get Default() As MethodCallingConvension
    Default = ccStdCall
End Property
Public Property Get APICall() As MethodCallingConvension
    APICall = ccStdCall
End Property

Public Property Get StdCall() As MethodCallingConvension
    StdCall
End Property
Public Property Get SysCall() As MethodCallingConvension
    Set SysCall = New MethodCallingConvension
    Call SysCall.Constructor0(CallingConvensionConstants.ccSysCall)
End Property
''<summary>
''  Specifies the default calling convention as determined by the common language
''  runtime. Use this calling convention for static methods. For instance or
''  virtual methods use HasThis (on CLR).
''</summary>
Public Property Get DotNetStandard() As MethodCallingConvension
    Set DotNetStandard = New MethodCallingConvension
    Call DotNetStandard.Constructor0(CallingConvensionConstants.ccDotNetStandard)
End Property
Public Property Get FastCall() As MethodCallingConvension
    Set FastCall = New MethodCallingConvension
    Call FastCall.Constructor0(CallingConvensionConstants.ccStdCall)
End Property
Public Property Get Cdecl_CC() As MethodCallingConvension
    Set Cdecl_CC = New MethodCallingConvension
    Call Cdecl_CC.Constructor0(CallingConvensionConstants.ccCdecl)
End Property
Public Property Get ThisCall() As MethodCallingConvension
    Set ThisCall = New MethodCallingConvension
    Call ThisCall.Constructor0(CallingConvensionConstants.ccThisCall)
End Property
''<summary>
''  Specifies that the signature is a function-pointer signature, representing
''  a call to an instance or virtual method (not a static method). If ExplicitThis
''  is set, HasThis must also be set. The first argument passed to the called
''  method is still a this pointer, but the type of the first argument is now
''  unknown. Therefore, a token that describes the type (or class) of the this
''  pointer is explicitly stored into its metadata signature.
''</summary>
Public Property Get ExplicitThis() As MethodCallingConvension
    Set ExplicitThis = New MethodCallingConvension
    Call ExplicitThis.Constructor0(CallingConvensionConstants.ccExplicitThis)
End Property
''<summary>
''  Specifies an instance or virtual method (not a static method). At run-time,
''  the called method is passed a pointer to the target object as its first argument
''  (the this pointer). The signature stored in metadata does not include the
''  type of this first argument, because the method is known and its owner class
''  can be discovered from metadata.
''</summary>
Public Property Get HasThis() As MethodCallingConvension
    Set HasThis = New MethodCallingConvension
    Call HasThis.Constructor0(CallingConvensionConstants.ccHasThis)
End Property
''<summary>
''  Specifies the calling convention for methods with variable arguments.<br>
''  It pushes address of first variable arguments list into stack.
''</summary>
Public Property Get VarArgs() As MethodCallingConvension
    Set VarArgs = New MethodCallingConvension
    Call VarArgs.Constructor0(CallingConvensionConstants.ccVarArgs)
End Property
''<summary>
''  Specifies that either the Standard or the VarArgs calling convention may
''  be used.
''</summary>
Public Property Get DotNetAny() As MethodCallingConvension
    Set DotNetAny = New MethodCallingConvension
    Call DotNetAny.Constructor0(CallingConvensionConstants.ccDotNetAny)
End Property

'*********************************************
'*********************************************
'
'</section>
'---------------------------------------------
'---------------------------------------------
