VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Exception"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "MintAPI exception type ,provides capsule of error data to be throwed by MintAPI.throw function."
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'@PROJECT_LICENSE

Option Base 0
Option Explicit
Const CLASSID As String = "Exception"
Const CLASSALIAS As String = APPLICATIONDOMAIN & "." & CLASSID
Const CLASSVERSIONSTRING As String = "0.1"
Const CLASSALIASSTRING As String = CLASSALIAS & "/" & CLASSVERSIONSTRING


Implements IAliasable


Dim mes As String
Dim errID As Long
Dim loc As String
Dim inner_Exception As Exception
Dim excType As ExceptionType

Dim bufErr As ErrObject

Friend Sub Initialize(Message As String, Location As String, ExceptionType As ExceptionType, ErrorID As Long, ByVal InnerException As Exception)
    mes = Message
    errID = ErrorID
    excType = ExceptionType
    loc = Location
    Set inner_Exception = InnerException
    Set bufErr = Err
End Sub

Friend Function ErrorObject() As ErrObject
    Set ErrorObject = bufErr
End Function

Public Property Get ExceptionType() As ExceptionType
    ExceptionType = excType
End Property
Friend Property Let ExceptionType(Value As ExceptionType)
    excType = Value
End Property
Public Property Get Message() As String
    Message = mes
End Property
Public Property Get ErrorID() As Long
    ErrorID = errID
End Property
Public Property Get Location() As String
    Location = loc
End Property
Public Property Get InnerException() As Exception
    Set InnerException = inner_Exception
End Property

Public Sub throw()
    Call mint_api_exceptions.throw(Me)
End Sub

Private Property Get IAliasable_Alias() As String
    IAliasable_Alias = CLASSALIASSTRING
End Property
