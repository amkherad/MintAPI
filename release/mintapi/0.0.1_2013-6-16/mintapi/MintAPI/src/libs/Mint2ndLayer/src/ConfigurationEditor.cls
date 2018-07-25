VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ConfigurationEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim frm As frmConfigurationEditor
Dim ParentForm

Private Sub Class_Initialize()
    Set frm = New frmConfigurationEditor
    Call frm.SetParent(Me)
End Sub

Friend Sub Initialize(Optional Path, Optional Parent)
    If VarType(Parent) = vbObject Then
        Set ParentForm = Parent
    Else
            ParentForm = Parent
    End If
End Sub

Public Sub Show()
    Call CheckNotModalState
    If IsEmpty(ParentForm) Then
        Call frm.Show(False)
    Else
        Call frm.Show(False, ParentForm)
    End If
End Sub
Public Sub ShowDialog()
    Call CheckNotModalState
    If IsEmpty(ParentForm) Then
        Call frm.Show(True)
    Else
        Call frm.Show(True, ParentForm)
    End If
End Sub
Public Sub Hide()
    Call frm.Hide
End Sub
Public Sub CloseForm()
    Call Unload(frm)
End Sub

Public Property Get Parent() As Long
    Parent = ParentForm
End Property
Public Property Let Parent(Value As Long)
    ParentForm = Value
End Property

Public Sub Focus()
    Call frm.SetFocus
End Sub

Public Sub Move(Left As Long, Optional Top, Optional Width, Optional Height)
    Call frm.Move(Left, Top, Width, Height)
End Sub

Public Property Get Left() As Long
    Left = frm.Left
End Property
Public Property Let Left(Value As Long)
    frm.Left = Value
End Property
Public Property Get Top() As Long
    Top = frm.Top
End Property
Public Property Let Top(Value As Long)
    frm.Top = Value
End Property
Public Property Get Width() As Long
    Width = frm.Width
End Property
Public Property Let Width(Value As Long)
    frm.Width = Value
End Property
Public Property Get Height() As Long
    Height = frm.Height
End Property
Public Property Let Height(Value As Long)
    frm.Height = Value
End Property

Public Property Get Visible() As Boolean
    Visible = frm.Visible
End Property
Public Property Let Visible(Value As Boolean)
    frm.Visible = Value
End Property
Public Sub Maximize()
    frm.WindowState = FormWindowStateConstants.vbMaximized
End Sub
Public Sub Minimize()
    frm.WindowState = FormWindowStateConstants.vbMinimized
End Sub
Public Sub Normal()
    frm.WindowState = FormWindowStateConstants.vbNormal
End Sub
Public Property Get WindowState() As FormWindowStateConstants
    WindowState = frm.WindowState
End Property
Public Property Let WindowState(Value As FormWindowStateConstants)
    frm.WindowState = Value
End Property

Private Sub Class_Terminate()
    Call Unload(frm)
    Set frm = Nothing
End Sub
