VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ControlCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Base 0
Option Explicit
Const CLASSID As String = "ControlCollection"
Const CLASSALIAS As String = APPLICATIONDOMAIN & "." & CLASSID
Const CLASSVERSIONSTRING As String = "0.1"
Const CLASSALIASSTRING As String = CLASSALIAS & "/" & CLASSVERSIONSTRING

Const SIGNAL_CONTROLADDED As String = "controladded"
Const SIGNAL_CONTROLREMOVED As String = "controlremoved"

Implements IAliasable
Implements ISupportSignal

Dim Col As New Collection

Dim IW As IWidget

Dim lastC As IWidget
Dim lastCAction As Long
Dim mySignalEmitter As New SignalEmitter


Private Sub Class_Initialize()
    Call mySignalEmitter.Initialize(Me)
End Sub
Friend Sub Initialize(Parent As IWidget)
    Set IW = Parent
End Sub

Public Sub Add(Component As IWidget)
    Call Col.Add(Component)
    lastCAction = Col.Count - 1
    Set lastC = Component
    Call mySignalEmitter.Emit(SIGNAL_CONTROLADDED)
End Sub
Public Sub Remove(Component As IWidget)
    Set lastC = Component
    Dim C, i As Long
    i = 1
    For Each C In Col
        If C Is Component Then
            Call Col.Remove(i)
            Call mySignalEmitter.Emit(SIGNAL_CONTROLREMOVED)
            lastCAction = i - 1
            Exit Sub
        End If
        i = i + 1
    Next
    throw ItemNotExistsException
End Sub
Public Sub RemoveByIndex(Index As Long)
    Set lastC = Item(Index)
    Call Col.Remove(Index + 1)
    Call mySignalEmitter.Emit(SIGNAL_CONTROLREMOVED)
End Sub

Public Property Get Item(Index As Long) As IWidget
    Set Item = Col(Index + 1)
End Property
Public Property Let Item(Index As Long, Value As IWidget)
    Set Col(Index + 1) = Value
End Property
Public Property Set Item(Index As Long, Value As IWidget)
    Set Col(Index + 1) = Value
End Property

Public Property Get Widget() As IWidget
    Set Widget = IW
End Property

Public Sub BringToFront(Index As Long)
    Call Item(Index).BringToFront
End Sub
Public Sub SendToBack(Index As Long)
    Call Item(Index).SendToBack
End Sub

Private Property Get IAliasable_Alias() As String
    IAliasable_Alias = CLASSALIASSTRING
End Property

Private Function ISupportSignal_Emitter() As SignalEmitter
    Set ISupportSignal_Emitter = mySignalEmitter
End Function
