VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IWidget"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Base 0
Option Explicit
Const CLASSID As String = "IWidget"
Const CLASSALIAS As String = APPLICATIONDOMAIN & "." & CLASSID
Const CLASSVERSIONSTRING As String = "0.1"
Const CLASSALIASSTRING As String = CLASSALIAS & "/" & CLASSVERSIONSTRING



Public Property Get Handle() As Long
    throw NotImplementedException
End Property
Public Property Get Name() As String
    throw NotImplementedException
End Property
Public Property Let Name(Value As String)
    throw NotImplementedException
End Property

Public Property Get Size() As MintAPI.Size
    throw NotImplementedException
End Property
Public Property Let Size(Value As MintAPI.Size)
    throw NotImplementedException
End Property

Public Property Get MinimumSize() As MintAPI.Size
    throw NotImplementedException
End Property
Public Property Let MinimumSize(Value As MintAPI.Size)
    throw NotImplementedException
End Property
Public Property Get MaximumSize() As MintAPI.Size
    throw NotImplementedException
End Property
Public Property Let MaximumSize(Value As MintAPI.Size)
    throw NotImplementedException
End Property

Public Property Get Position() As MintAPI.Position
    throw NotImplementedException
End Property
Public Property Let Position(Value As MintAPI.Position)
    throw NotImplementedException
End Property

Public Property Get Visible() As Boolean
    throw NotImplementedException
End Property
Public Property Let Visible(Value As Boolean)
    throw NotImplementedException
End Property

Public Property Get Enabled() As Boolean
    throw NotImplementedException
End Property
Public Property Let Enabled(Value As Boolean)
    throw NotImplementedException
End Property

'MUST CHANGE TO REGION ### THIS MEAN GRAPHICAL REGION.
Public Property Get Region() As MintAPI.Rectangle
    throw NotImplementedException
End Property
Public Property Let Region(Value As MintAPI.Rectangle)
    throw NotImplementedException
End Property

Public Sub BringToFront()
    throw NotImplementedException
End Sub
Public Sub SendToBack()
    throw NotImplementedException
End Sub

Public Sub Move(Left As Long, Top As Long)
    throw NotImplementedException
End Sub
Public Sub Resize(Width As Long, Height As Long)
    throw NotImplementedException
End Sub
Public Sub SetRectangle(Left As Long, Top As Long, Width As Long, Height As Long)
    throw NotImplementedException
End Sub
Public Function GetRectangle() As Rectangle
    throw NotImplementedException
End Function

Public Property Get Font() As MintAPI.Font
    throw NotImplementedException
End Property
Public Property Let Font(Value As MintAPI.Font)
    throw NotImplementedException
End Property

Public Property Get ForeColor() As Color
    throw NotImplementedException
End Property
Public Property Let ForeColor(Value As Color)
    throw NotImplementedException
End Property
Public Property Get BackColor() As Color
    throw NotImplementedException
End Property
Public Property Let BackColor(Value As Color)
    throw NotImplementedException
End Property

Public Property Get Parent() As IWidget
    throw NotImplementedException
End Property

Public Function Emitter() As SignalEmitter
    throw NotImplementedException
End Function
