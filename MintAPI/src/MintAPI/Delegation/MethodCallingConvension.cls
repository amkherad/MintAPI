VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MethodCallingConvension"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"MetaObject"
'@PROJECT_LICENSE

''<class name="MethodCallingConvension" tags="">
''  <summary>
''    Some class description.
''  </summary>
''  <constructors>
''      <constructor name="" friend>
''          Some constructor description.
''      </constructor>
''  </constructors>
''  <using>Some class alias.</using>
''  <idea>Me</idea>
''  <sample></sample>
''</class>

Option Base 0
Option Explicit

Implements IObject
Implements ICloneable

'=============================================
'=============================================
'=============================================
'<section Meta Data>
'
Const CLASSID As String = "MethodCallingConvension" '#DO NOT REMOVE IT
Const CLASSALIAS As String = APPLICATIONDOMAIN & "." & CLASSID '#DO NOT REMOVE IT
Const CLASSVERSIONSTRING As String = "0.1" '#DO NOT REMOVE IT
Const CLASSALIASSTRING As String = CLASSALIAS & "/" & CLASSVERSIONSTRING '#DO NOT REMOVE IT

Const SIGNAL_TERMINATING As String = "terminating"
Const SLOT_METAOBJECT As String = "metaobject"
'
'</section>
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section API/Type Declarations>
'
Public Enum CallingConvensionPushOrder
    ccpoRightToLeft = &H0
    ccpoLeftToRight = &H8
    ccpo_EnumMask = &H8
End Enum
Public Enum CallingConvensionFlags
    ccfLeftToRightArgumentPush = ccpoLeftToRight
    ccfRightToLeftArgumentPush = ccpoRightToLeft
    ccfTreatRetvalAsLastArgument = &H0
    ccfPushRetvalAtLast = &H0
    ccfPushRetvalAtFirst = &H4
    ccfPopArgumentsAfterCall = &H2
    ccfPopRetvalAfterCall = &H1
    ccfCleanStackAfterCall = ccfPopArgumentsAfterCall Or ccfPopRetvalAfterCall
    ccfCleanStackByCallee = ccfPopRetvalAfterCall
    ccfCleanStackByCaller = ccfCleanStackAfterCall
    
    ccfUseDllGetLastError = &H10
    ccfUseStandardFramedErrorHandling = &H20
    
    ccfUserDefinedCallerToCallMethod = &H100
    ccfCallStackPusherToPushArguments = &H200
    ccfCallGetRetvalToPopRetval = &H400
    ccfCallCleanerToPopStack = &H800
    
    ccfEstablishThreadSafeFoundation = &H8000
End Enum
Public Enum CallingConvensionConstants
    ccStdCall = ccpoRightToLeft Or ccfCleanStackByCallee
    ccCdecl = ccpoRightToLeft Or ccfCleanStackByCaller
    ccSysCall = 3
    ccFastCall = 2
    ccThisCall = 3 'Set ECX = Class ptr.
    ccPascal = 4
    ccMint = 5
    ccDotNetStandard
    ccExplicitThis
    ccHasThis
    ccVarArgs
    ccDotNetAny
    ccCustom = 99
End Enum
'
'</section>
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section Variables/Events>
'
Dim This_Constructed As Boolean
Dim myMetaObject As MetaObject

'Dim p_MethodWeak As New WeakReference
Dim p_Flags As CallingConvensionFlags
'
'</section>
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section Constructors>
'
''<summary>
''  Basic constructor to use by MintAPI class provider,
''  When creating new instance of this class.
''</summary>
''<params>
''  <param name="Arguments">
''      A list of construction arguments.
''  </param>
''</params>
''<cargs order="0" count="1">
''  <carg type="long" name="ID"></carg>
''</cargs>
Friend Sub Constructor(ByVal CInfo As ObjectConstructionInfo)
    'Set p_MethodWeak.Object = Arguments(0)
    
    This_Constructed = True
End Sub
Friend Sub Constructor0(ByVal Flags As CallingConvensionFlags)
    p_Flags = Flags
    
    This_Constructed = True
End Sub
Friend Sub ConstructorCopy(ByVal That As MethodCallingConvension)
    p_Flags = That.Flags
    
    This_Constructed = True
End Sub
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

Public Property Get ArgumentPushOrder() As CallingConvensionPushOrder
    ArgumentPushOrder = (p_Flags And ccpo_EnumMask)
End Property
Public Property Let ArgumentPushOrder(ByVal Value As CallingConvensionPushOrder)
    Select Case Value
        Case ccpoLeftToRight, ccpoRightToLeft
        Case Else
            throw Exps.InvalidArgumentException
    End Select
    
    p_Flags = p_Flags Or Value
End Property

Public Function SetFlags(ByVal Flag As CallingConvensionFlags) As Method
    p_Flags = p_Flags Or Flag
End Function

Public Property Get Flags() As CallingConvensionConstants
    Flags = p_Flags
End Property

'*********************************************
'*********************************************
'
'</section>
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section Basic Members>
'
Friend Function EqualsFlags(ByVal Flags As CallingConvensionFlags) As Boolean
    EqualsFlags = (Flags = p_Flags)
End Function
Public Function Equals(Value As Variant) As Boolean
    If IsObject(Value) Then
        If Value Is Nothing Then Exit Function
        
        Dim MCC As MethodCallingConvension
        Set MCC = Value
        If TypeOf MCC Is MethodCallingConvension Then
            Equals = (p_Flags = MCC.Flags)
        End If
    End If
End Function
Public Function ToString() As String
    ToString = CLASSALIAS
End Function
Public Function GetHashCode() As Long
    GetHashCode = Objects.GetHashCode(Me)
End Function
Public Function MetaObject() As MetaObject
    If myMetaObject Is Nothing Then
        Call Objects.MetaObject.Setup(myMetaObject, Me, True, CLASSALIASSTRING)
    End If
    Set MetaObject = myMetaObject
End Function

Public Function Clone() As MethodCallingConvension
    Set Clone = New MethodCallingConvension
    Call Clone.ConstructorCopy(Me)
End Function
'
'</section>
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section Private helpers>
'
'Private Sub Class_InitProperties()
'
'End Sub
Private Sub Class_Terminate()
    If Not myMetaObject Is Nothing Then
        Signal.Emit Me, SIGNAL_TERMINATING
        Call myMetaObject.Dispose
        Set myMetaObject = Nothing
    End If
End Sub
'
'</section>
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section IObject Implementation>
'
Private Sub IObject_Construct(ByVal CInfo As ObjectConstructionInfo): Call Constructor(CInfo): End Sub
Private Function IObject_Equals(Value As Variant) As Boolean: IObject_Equals = Equals(Value): End Function
Private Function IObject_GetHashCode() As Long: IObject_GetHashCode = GetHashCode: End Function
Private Function IObject_MetaObject() As MetaObject: Set IObject_MetaObject = MetaObject: End Function
Private Function IObject_ToString() As String: IObject_ToString = ToString: End Function
Private Function IObject_Stub(ByVal ProxyInfo As ObjectProxyInfo) As Boolean
    
    IObject_Stub = False
End Function
'
'</section>
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section ICloneable Implementations>
'
Private Function ICloneable_Clone() As Object: Set ICloneable_Clone = Clone: End Function
'
'</section>
'---------------------------------------------
'---------------------------------------------
