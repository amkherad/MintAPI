VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SourceCodeManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Event QueryUnload(Cancel As Boolean)

Dim frm As frmSourceCodeManager
Dim ParentForm
Dim Disposed As Boolean

Private Sub Class_Initialize()
    Set frm = New frmSourceCodeManager
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
    If IsDisposed Then throw Exceptions.DisposedException
    Call CheckNotModalState
    If IsEmpty(ParentForm) Then
        Call frm.Show(False)
    Else
        Call frm.Show(False, ParentForm)
    End If
End Sub
Public Sub ShowDialog()
    If IsDisposed Then throw Exceptions.DisposedException
    Call CheckNotModalState
    If IsEmpty(ParentForm) Then
        Call frm.Show(True)
    Else
        Call frm.Show(True, ParentForm)
    End If
End Sub
Public Sub Hide()
    If IsDisposed Then throw Exceptions.DisposedException
    Call frm.Hide
End Sub
Public Sub CloseForm()
    If IsDisposed Then Exit Sub
    Call Unload(frm)
    Set frm = Nothing
    Disposed = True
End Sub
Friend Sub ClosingForm(Cancel As Boolean)
    RaiseEvent QueryUnload(Cancel)
    If Cancel Then Exit Sub
    Set frm = Nothing
    Disposed = True
End Sub

Public Property Get IsDisposed() As Boolean
    IsDisposed = Disposed Or (frm Is Nothing)
End Property

Public Property Get Parent() As Long
    Parent = ParentForm
End Property
Public Property Let Parent(Value As Long)
    ParentForm = Value
End Property

Public Sub Focus()
    If IsDisposed Then throw Exceptions.DisposedException
    Call frm.SetFocus
End Sub

Public Sub Move(Left As Long, Optional Top, Optional Width, Optional Height)
    If IsDisposed Then throw Exceptions.DisposedException
    Call frm.Move(Left, Top, Width, Height)
End Sub

Public Property Get Text() As String
    If IsDisposed Then throw Exceptions.DisposedException
    Text = frm.Caption
End Property
Public Property Let Text(Value As String)
    If IsDisposed Then throw Exceptions.DisposedException
    frm.Caption = Value
End Property

Public Property Get Left() As Long
    If IsDisposed Then throw Exceptions.DisposedException
    Left = frm.Left
End Property
Public Property Let Left(Value As Long)
    If IsDisposed Then throw Exceptions.DisposedException
    frm.Left = Value
End Property
Public Property Get Top() As Long
    If IsDisposed Then throw Exceptions.DisposedException
    Top = frm.Top
End Property
Public Property Let Top(Value As Long)
    If IsDisposed Then throw Exceptions.DisposedException
    frm.Top = Value
End Property
Public Property Get Width() As Long
    If IsDisposed Then throw Exceptions.DisposedException
    Width = frm.Width
End Property
Public Property Let Width(Value As Long)
    If IsDisposed Then throw Exceptions.DisposedException
    frm.Width = Value
End Property
Public Property Get Height() As Long
    If IsDisposed Then throw Exceptions.DisposedException
    Height = frm.Height
End Property
Public Property Let Height(Value As Long)
    If IsDisposed Then throw Exceptions.DisposedException
    frm.Height = Value
End Property

Public Property Get Visible() As Boolean
    If IsDisposed Then throw Exceptions.DisposedException
    Visible = frm.Visible
End Property
Public Property Let Visible(Value As Boolean)
    If IsDisposed Then throw Exceptions.DisposedException
    frm.Visible = Value
End Property
Public Sub Maximize()
    If IsDisposed Then throw Exceptions.DisposedException
    frm.WindowState = FormWindowStateConstants.vbMaximized
End Sub
Public Sub Minimize()
    If IsDisposed Then throw Exceptions.DisposedException
    frm.WindowState = FormWindowStateConstants.vbMinimized
End Sub
Public Sub Normal()
    If IsDisposed Then throw Exceptions.DisposedException
    frm.WindowState = FormWindowStateConstants.vbNormal
End Sub
Public Property Get WindowState() As FormWindowStateConstants
    If IsDisposed Then throw Exceptions.DisposedException
    WindowState = frm.WindowState
End Property
Public Property Let WindowState(Value As FormWindowStateConstants)
    If IsDisposed Then throw Exceptions.DisposedException
    frm.WindowState = Value
End Property

Private Sub Class_Terminate()
    If IsDisposed Then Exit Sub
    Call Unload(frm)
    Set frm = Nothing
End Sub

