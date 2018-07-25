VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MenuItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'@PROJECT_LICENSE

''<class name="MenuItem" tags="">
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
'Implements ICloneable
Implements IDisposable

'=============================================
'=============================================
'=============================================
'<section Meta Data>
'
Const CLASSID As String = "MenuItem" '#DO NOT REMOVE IT
Const CLASSALIAS As String = APPLICATIONDOMAIN & "." & CLASSID '#DO NOT REMOVE IT
Const CLASSVERSIONSTRING As String = "0.1" '#DO NOT REMOVE IT
Const CLASSALIASSTRING As String = CLASSALIAS & "/" & CLASSVERSIONSTRING '#DO NOT REMOVE IT

Const SIGNAL_TERMINATING                As String = "terminating"
Const SLOT_METAOBJECT                   As String = "metaobject"
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
Public Event Click(ByVal E As EventArgs)
Public Event ChildClick(ByVal E As MenuItemEventArgs)
Public Event Opening(ByVal E As CancelEventArgs)
'Public Event Closed(ByVal E As EventArgs)

Dim This_Constructed As Boolean
Dim myMetaObject As MetaObject

Private p_VBInstance                        As VBInstanceWrapper
Private WithEvents p_Menu_Handler           As CommandBarEvents
Attribute p_Menu_Handler.VB_VarHelpID = -1
Private p_Menu                              As CommandBarControl
Private p_Caption                           As String

Private WithEvents p_Children               As MenuItemChildren
Attribute p_Children.VB_VarHelpID = -1
Private p_WeakParent                        As New WeakReference

Private p_IsParent                          As Boolean
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
Friend Sub Constructor0(ByVal Wrapper As VBInstanceWrapper, ByVal Parent As MenuItem, ByVal mnu As CommandBarControl)
    Set p_VBInstance = Wrapper
    Set p_WeakParent.Object = Parent
    Set p_Menu = mnu
    Set p_Menu_Handler = p_VBInstance.VBE.Events.CommandBarEvents(mnu)
    
    This_Constructed = True
End Sub
'Friend Sub ConstructorCopy(ByVal That As {My Object Type})
'
'    This_Constructed = True
'End Sub
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

Public Property Get Instance() As VBInstanceWrapper
    Set Instance = p_VBInstance
End Property

Public Property Get Menu() As CommandBarControl
    Set Menu = p_Menu
End Property

Public Property Get MenuEventHandler() As CommandBarEvents
    Set MenuEventHandler = p_Menu_Handler
End Property

Public Property Get Children() As MenuItemChildren
    Set Children = p_Children
End Property

Public Property Get Caption() As String
    If p_Menu Is Nothing Then
        Caption = p_Caption
    Else
        Caption = p_Menu.Caption
    End If
End Property
Public Property Let Caption(ByVal Value As String)
    If Not p_Menu Is Nothing Then _
        p_Menu.Caption = Value
    p_Caption = Value
End Property

Friend Property Let IsParent(ByVal Value As Boolean)
    p_IsParent = Value
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
    If myMetaObject Is Nothing Then
        Call Objects.MetaObject.Setup(myMetaObject, Me, True, CLASSALIASSTRING)
    End If
    Set MetaObject = myMetaObject
End Function

'Public Function Clone() As Object
'
'End Function

Public Sub Dispose()
    Call p_Children.Dispose
    If Not p_Menu Is Nothing Then
        Call p_Menu.Delete
        Set p_Menu = Nothing
    End If
    Set p_WeakParent.Object = Nothing
End Sub
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
    Set p_Children = New MenuItemChildren
    Call p_Children.Constructor0(p_VBInstance, Me)
End Sub
Private Sub Class_Terminate()
    If Not myMetaObject Is Nothing Then
        Signal.Emit Me, SIGNAL_TERMINATING
        Call myMetaObject.Dispose
        Set myMetaObject = Nothing
    End If
    
    Call Dispose
End Sub

Friend Sub SetParent(ByVal MenuItem As MenuItem)
    If Not MenuItem Is Nothing Then
        If Not p_Menu Is Nothing Then _
            Call p_Menu.Delete
        
        Set p_VBInstance = MenuItem.Instance
        Set p_WeakParent.Object = MenuItem
        
        Dim Obj As Object
        Set Obj = MenuItem.Menu
        
        If p_IsParent Then
            Set p_Menu = Obj.Controls.Add(10)
        Else
            Set p_Menu = Obj.Controls.Add(1)
        End If
        p_Menu.Caption = p_Caption
        Set p_Menu_Handler = p_VBInstance.VBE.Events.CommandBarEvents(p_Menu)
    Else
        Set p_WeakParent.Object = Nothing
        Call p_Menu.Delete
        Set p_Menu = Nothing
        Set p_Menu_Handler = Nothing
    End If
End Sub

Public Sub OnClick(ByVal E As MenuItemEventArgs)
    RaiseEvent ChildClick(E.Clone)
    
    If p_WeakParent.HasValue Then
        Dim P As MenuItem
        Set P = p_WeakParent.Object
        Call P.OnClick(E)
    End If
End Sub

Private Sub p_Menu_Handler_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    If p_Children.Count = 0 Then
        RaiseEvent Click(EventArgs(Me))

        If p_WeakParent.HasValue Then
            Dim P As MenuItem
            Set P = p_WeakParent.Object
            Call P.OnClick(MenuItemEventArgs(Me, Me))
        End If
    Else
        Dim CE As CancelEventArgs
        Set CE = CancelEventArgs(Me, False)
        RaiseEvent Opening(CE)
        If CE.Cancel Then CancelDefault = True: Exit Sub
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
            throw Exps.UnknownNotificationException, Me
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
'Private Function ICloneable_Clone() As Object
'    Set ICloneable_Clone = Clone
'End Function
'
'</section>
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section IDisposable Implementations>
'
Private Sub IDisposable_Dispose(): Call Dispose: End Sub
'
'</section>
'---------------------------------------------
'---------------------------------------------
