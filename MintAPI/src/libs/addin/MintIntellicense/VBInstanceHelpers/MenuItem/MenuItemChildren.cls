VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MenuItemChildren"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'@PROJECT_LICENSE

''<class name="MenuItemChildren" tags="">
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
Implements IEnumerable
Implements ICollection
Implements IDisposable

'=============================================
'=============================================
'=============================================
'<section Meta Data>
'
Const CLASSID As String = "MenuItemChildren" '#DO NOT REMOVE IT
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
Dim This_Constructed As Boolean
Dim myMetaObject As MetaObject

Private p_WeakParent As New WeakReference
Private p_VBInstance As VBInstanceWrapper

Public Event ItemAdded(ByVal E As MenuItemEventArgs)
Public Event ItemRemoved(ByVal E As MenuItemEventArgs)

Public Event ItemsChanged(ByVal E As EventArgs)

Private p_Items As New List
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
Friend Sub Constructor0(ByVal Wrapper As VBInstanceWrapper, ByVal Menu As MenuItem)
    Set p_VBInstance = Wrapper
    Set p_WeakParent.Object = Menu
    
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

Private Function GetEnumerator() As MintAPI.IEnumerator
    If Not p_WeakParent.IsAlive Then throw Exps.InvalidStatusException
    Set GetEnumerator = p_Items.GetEnumerator
End Function
Private Function NewEnum() As stdole.IUnknown
    If Not p_WeakParent.IsAlive Then throw Exps.InvalidStatusException
    Set NewEnum = Enumerator.GetNewEnum(p_Items)
End Function

Public Property Get Parent() As MenuItem
    If Not p_WeakParent.IsAlive Then throw Exps.InvalidStatusException
    Set Parent = p_WeakParent.Object
End Property

Public Sub Add(ByVal MenuItem As MenuItem)
    If MenuItem Is Nothing Then throw Exps.ArgumentNullException
    If Not p_WeakParent.IsAlive Then throw Exps.InvalidStatusException
    
    Call p_Items.Append(MenuItem)
    
    Call MenuItem.SetParent(p_WeakParent.Object)
    
    RaiseEvent ItemAdded(MenuItemEventArgs(Me, MenuItem))
    RaiseEvent ItemsChanged(EventArgs(Me))
End Sub
Public Sub Insert(ByVal MenuItem As MenuItem, ByVal Index As Long)
    If MenuItem Is Nothing Then throw Exps.ArgumentNullException
    If Not p_WeakParent.IsAlive Then throw Exps.InvalidStatusException
    Call p_Items.Insert(Index, MenuItem)
    
    Call MenuItem.SetParent(p_WeakParent.Object)
    
    RaiseEvent ItemAdded(MenuItemEventArgs(Me, MenuItem))
    RaiseEvent ItemsChanged(EventArgs(Me))
End Sub
Public Sub Remove(ByVal MenuItem As MenuItem)
    If MenuItem Is Nothing Then throw Exps.ArgumentNullException
    If Not p_WeakParent.IsAlive Then throw Exps.InvalidStatusException
    Call p_Items.Remove(MenuItem)
    
    Call MenuItem.SetParent(Nothing)
    
    RaiseEvent ItemRemoved(MenuItemEventArgs(Me, MenuItem))
    RaiseEvent ItemsChanged(EventArgs(Me))
End Sub

Public Property Get Items() As ReadOnlyVector
    If Not p_WeakParent.IsAlive Then throw Exps.InvalidStatusException
    Set Items = p_Items.AsReadOnly
End Property

Public Property Get Item(ByVal Index As Long) As MenuItem
Attribute Item.VB_ProcData.VB_Invoke_Property = ";List"
Attribute Item.VB_UserMemId = 0
    If Not p_WeakParent.IsAlive Then throw Exps.InvalidStatusException
    Set Item = p_Items.Item(Index)
End Property

Public Property Get Count() As Long
    If Not p_WeakParent.IsAlive Then throw Exps.InvalidStatusException
    Count = p_Items.Count
End Property

Public Sub CopyTo(Arr As Variant, ByVal Index As Long)
    If Not p_WeakParent.IsAlive Then throw Exps.InvalidStatusException
    Call p_Items.CopyTo(Arr, Index)
End Sub

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
    Call p_Items.Clear
    Dim Obj As MenuItem
    For Each Obj In p_Items
        Call Obj.Dispose
    Next
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
'Private Sub Class_Initialize()
'    ...
'End Sub
Private Sub Class_Terminate()
    If Not myMetaObject Is Nothing Then
        Signal.Emit Me, SIGNAL_TERMINATING
        Call myMetaObject.Dispose
        Set myMetaObject = Nothing
    End If
    
    Call Dispose
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
'<section ICloneable Implementations>
'
Private Function IEnumerable_GetEnumerator() As MintAPI.IEnumerator
    Set IEnumerable_GetEnumerator = GetEnumerator
End Function
Private Function IEnumerable_NewEnum() As stdole.IUnknown
    Set IEnumerable_NewEnum = NewEnum
End Function
'
'</section>
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section ICollection Implementations>
'
Private Sub ICollection_CopyTo(Arr As Variant, ByVal Index As Long): Call CopyTo(Arr, Index): End Sub
Private Property Get ICollection_Count() As Long: ICollection_Count = Count: End Property
Private Function ICollection_GetEnumerator() As MintAPI.IEnumerator: Set ICollection_GetEnumerator = GetEnumerator: End Function
Private Function ICollection_NewEnum() As stdole.IUnknown: Set ICollection_NewEnum = NewEnum: End Function
'
'</section>
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section ICollection Implementations>
'
Private Sub IDisposable_Dispose(): Call Dispose: End Sub
'
'</section>
'---------------------------------------------
'---------------------------------------------
