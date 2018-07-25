VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MediaGenericAttributes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'@PROJECT_LICENSE

''<class name="MediaGenericAttributes" tags="">
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
Const CLASSID As String = "MediaGenericAttributes" '#DO NOT REMOVE IT
Const CLASSALIAS As String = APPLICATIONDOMAIN & "." & CLASSID '#DO NOT REMOVE IT
Const CLASSVERSIONSTRING As String = "0.1" '#DO NOT REMOVE IT
Const CLASSALIASSTRING As String = CLASSALIAS & "/" & CLASSVERSIONSTRING '#DO NOT REMOVE IT

Const SIGNAL_TERMINATING                As String = "terminating"
Const SLOT_METAOBJECT                   As String = "metaobject"

Const SIGNAL_VOLUMECHANGED              As String = "volumechanged"
Const SIGNAL_BALANCECHANGED             As String = "balancechanged"
Const SIGNAL_TREBLECHANGED              As String = "treblechanged"
Const SIGNAL_BASSCHANGED                As String = "basschanged"

Const QUERYINFORMATION_ALL              As String = "all"
Const QUERYINFORMATION_DEBUGINFO        As String = "dbginfo"
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
'<section Variables/Events>
'
Dim This_Constructed As Boolean
Dim myMetaObject As MetaObject

Public Event VolumeChanged(ByVal E As EventArgs)
Public Event BalanceChanged(ByVal E As EventArgs)
Public Event TrebleChanged(ByVal E As EventArgs)
Public Event BassChanged(ByVal E As EventArgs)

Dim p_Volume_Support As Boolean, p_Volume As Long
Dim p_LeftBalance_Support As Boolean, p_LeftBalance As Long
Dim p_RightBalance_Support As Boolean, p_RightBalance As Long

Dim p_Treble_Support As Boolean, p_Treble As Long
Dim p_Bass_Support As Boolean, p_Bass As Long
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
Friend Sub ConstructorArg(ByVal Arguments As ArgumentList)
    
    This_Constructed = True
End Sub
Friend Sub Constructor0(ByVal SupportVolume As Boolean, ByVal SupportLeftBalance As Boolean, _
    ByVal SupportRightBalance As Boolean)
    
    p_Volume_Support = SupportVolume
    p_LeftBalance_Support = SupportLeftBalance
    p_RightBalance_Support = SupportRightBalance
    
    This_Constructed = True
End Sub
Friend Sub ConstructorCopy(ByVal That As MediaGenericAttributes)
    
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

Public Property Get Volume() As Long
    If Not p_Volume_Support Then throw Exps.NotSupportedException
    Volume = p_Volume
End Property
Public Property Let Volume(ByVal Value As Long)
    If Not p_Volume_Support Then throw Exps.NotSupportedException
    If (Value < 0) Or (Value > 100) Then throw Exps.OutOfRangeException
    p_Volume = Value
    Signal.Emit Me, SIGNAL_VOLUMECHANGED
    RaiseEvent VolumeChanged(EventArgs(Me))
End Property

Public Property Get LeftBalance() As Long
    If Not p_LeftBalance_Support Then throw Exps.NotSupportedException
    LeftBalance = p_LeftBalance
End Property
Public Property Let LeftBalance(ByVal Value As Long)
    If Not p_LeftBalance_Support Then throw Exps.NotSupportedException
    If (Value < -100) Or (Value > 100) Then throw Exps.OutOfRangeException
    p_LeftBalance = Value
    Signal.Emit Me, SIGNAL_BALANCECHANGED
    RaiseEvent BalanceChanged(EventArgs(Me))
End Property
Public Property Get RightBalance() As Long
    If Not p_RightBalance_Support Then throw Exps.NotSupportedException
    RightBalance = p_RightBalance
End Property
Public Property Let RightBalance(ByVal Value As Long)
    If Not p_RightBalance_Support Then throw Exps.NotSupportedException
    If (Value < -100) Or (Value > 100) Then throw Exps.OutOfRangeException
    p_RightBalance = Value
    Signal.Emit Me, SIGNAL_BALANCECHANGED
    RaiseEvent BalanceChanged(EventArgs(Me))
End Property

Public Property Get Balance() As Long
    If Not (p_RightBalance_Support And p_LeftBalance_Support) Then throw Exps.NotSupportedException
    Balance = (p_RightBalance - p_LeftBalance)
End Property
Public Property Let Balance(ByVal Value As Long)
    If Not (p_RightBalance_Support And p_LeftBalance_Support) Then throw Exps.NotSupportedException
    If (Value < -100) Or (Value > 100) Then throw Exps.OutOfRangeException
    If Value < 0 Then
        p_RightBalance = 100 - Abs(Value)
        p_LeftBalance = 100
    Else
        p_RightBalance = 100
        p_LeftBalance = 100 - Value
    End If
    Signal.Emit Me, SIGNAL_BALANCECHANGED
    RaiseEvent BalanceChanged(EventArgs(Me))
End Property

Public Property Get Treble() As Long
    If Not p_Treble_Support Then throw Exps.NotSupportedException
    Treble = p_Treble
End Property
Public Property Let Treble(ByVal Value As Long)
    If Not p_Treble_Support Then throw Exps.NotSupportedException
    If (Value < 0) Or (Value > 100) Then throw Exps.OutOfRangeException
    p_Treble = Value
    Signal.Emit Me, SIGNAL_TREBLECHANGED
    RaiseEvent TrebleChanged(EventArgs(Me))
End Property

Public Property Get Bass() As Long
    If Not p_Bass_Support Then throw Exps.NotSupportedException
    Bass = p_Bass
End Property
Public Property Let Bass(ByVal Value As Long)
    If Not p_Bass_Support Then throw Exps.NotSupportedException
    If (Value < 0) Or (Value > 100) Then throw Exps.OutOfRangeException
    p_Bass = Value
    Signal.Emit Me, SIGNAL_BASSCHANGED
    RaiseEvent BassChanged(EventArgs(Me))
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
Public Function Equals(Value As Variant) As Boolean
    Equals = Objects.Equals(Me, Value)
End Function
Public Function ToString() As String
    ToString = CLASSALIAS
End Function
Public Function GetHashCode() As Long
    GetHashCode = Objects.GetHashCode(Me)
End Function
Public Function MetaObject() As MetaObject
    If Not myMetaObject Is Nothing Then
        Set MetaObject = myMetaObject
    End If
    Call Objects.MetaObject.Setup(myMetaObject, Me, True, CLASSALIASSTRING)
    Set MetaObject = myMetaObject
End Function

Public Function Clone() As MediaGenericAttributes
    Set Clone = New MediaGenericAttributes
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
Private Sub Class_Initialize()
    p_Volume = 100 '%
    p_Volume_Support = True
    p_LeftBalance_Support = True: p_RightBalance_Support = True
    p_Treble_Support = True: p_Bass_Support = True
End Sub
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
Private Sub IObject_Initialize(ByVal Arguments As ArgumentList)
    Call ConstructorArg(Arguments)
End Sub
Private Function IObject_Equals(Value As Variant) As Boolean
    IObject_Equals = Equals(Value)
End Function
Private Function IObject_GetHashCode() As Long
    IObject_GetHashCode = GetHashCode
End Function
Private Function IObject_MetaObject() As MetaObject
    Set IObject_MetaObject = MetaObject
End Function
Private Function IObject_ToString() As String
    IObject_ToString = ToString
End Function
Private Function IObject_QueryInformation(ByVal OQR As ObjectQR) As Boolean
    
    IObject_QueryInformation = False
End Function
Private Function IObject_Notify(ByVal Name As String, ByVal Arguments As ArgumentList) As Variant
    Select Case LCase$(Name)
        'Case ""
            
        Case Else
            throw Exps.UnknownNotificationException
    End Select
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
Private Function ICloneable_Clone() As Object
    Set ICloneable_Clone = Clone
End Function
'
'</section>
'---------------------------------------------
'---------------------------------------------
