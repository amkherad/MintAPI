VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ExceptionMethods"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@PROJECT_LICENSE

Option Base 0
Option Explicit
Const CLASSID As String = "ExceptionMethods"
Const CLASSALIAS As String = APPLICATIONDOMAIN & "." & CLASSID
Const CLASSVERSIONSTRING As String = "0.1"
Const CLASSALIASSTRING As String = CLASSALIAS & "/" & CLASSVERSIONSTRING


Implements IAliasable


Friend Sub Initialize(Optional Arguments)
    
End Sub

Public Sub throw0(ByRef LibraryName As String, ByRef ModuleName As String, ByVal SourceMethodName As String, ByVal ExceptionDescription As String, Optional ByVal Arguments As String = "", Optional ByVal ErrorNumber As Long, Optional ByVal ExceptionType As Long, Optional HelpFile, Optional HelpContext)
    Call mint_api_exceptions.throw0(LibraryName, ModuleName, SourceMethodName, ExceptionDescription, Arguments, ErrorNumber, ExceptionType, HelpFile, HelpContext)
End Sub
Public Sub throw(Exception As Exception)
    Call mint_api_exceptions.throw(Exception)
End Sub
Public Sub rethrow()
    Call mint_api_exceptions.rethrow
End Sub
Public Sub throwD(Description As String, Optional Location As String = ""): Call mint_api_exceptions.throwD(Description, Location): End Sub
Public Sub throwN(): Call mint_api_exceptions.throwN: End Sub

Public Sub Ensure(Condition As Boolean, Exception As Exception)
    If Not Condition Then throw Exception
End Sub
Public Sub Assert(Condition As Boolean, Exception As Exception)
    If Not Condition Then throw Exception
End Sub
Public Sub AssertM(Condition As Boolean, Message As String, Optional Location As String)
    If Not Condition Then throw Exception(Message, Location)
End Sub


Public Function Location(Library As String, Module As String, Func As String, Optional Arguments As String = "") As String
    Location = mint_api_exceptions.Location(Library, Module, Func, Arguments)
End Function


Private Property Get IAliasable_Alias() As String
    IAliasable_Alias = CLASSALIASSTRING
End Property
