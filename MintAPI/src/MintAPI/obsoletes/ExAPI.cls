VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ExAPI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'@PROJECT_LICENSE


'Ideas from Me , VBA Automation

Option Base 0
Option Explicit
Const CLASSID As String = "ExAPI"
Const CLASSALIAS As String = APPLICATIONDOMAIN & "." & CLASSID
Const CLASSVERSIONSTRING As String = "0.1"
Const CLASSALIASSTRING As String = CLASSALIAS & "/" & CLASSVERSIONSTRING


Implements IAliasable


Dim Args

Friend Sub Initialize(Optional Arguments)
    If IsMissing(Arguments) Then Exit Sub
    If VarType(Arguments) = VBObject Then
        Set Args = Arguments
    Else
            Args = Arguments
    End If
End Sub

Public Function BeginExecutionOf( _
    ExecObj As Object, MethodName As String, _
    Args As ArgumentList, Optional ByVal IsCocurrent As Boolean = False) As Variant
    
End Function
Public Function BreakExecution()
    
End Function
Public Sub SetArgument( _
    Object As Object, ArgumentName As String, Value As Variant)
    
End Sub
Public Sub SetArguments( _
    Object As Object, ParamArray Args())
    
End Sub
Public Sub SetArgumentsFromList( _
    Object As Object, ArgList As ArgumentList) 'IMPORTANT : All argument list's arguments must have a name.
    
End Sub

Public Function KnownClassType(Object As Object) As String
    
End Function

Public Function CreateClass(ClassName As String, Properties() As Argument) As Object
    
End Function

'
'Public Function CreateDownloader() As Object
'    Dim props() As Argument
'    'Initialize the props
'    Set CreateDownloader = CreateClass("mint_downloader", props)
'End Function
'Public Function CreateService() As Object
'    Dim props() As Argument
'    'Initialize the props
'    Set CreateDownloader = CreateClass("mint_service", props)
'End Function
'Public Function CreateHacker() As Object
'    Dim props() As Argument
'    'Initialize the props
'    Set CreateDownloader = CreateClass("mint_hacker", props)
'End Function
'Public Function CreateCompresser() As Object
'    Dim props() As Argument
'    'Initialize the props
'    Set CreateDownloader = CreateClass("mint_compresser", props)
'End Function

Private Property Get IAliasable_Alias() As String
    IAliasable_Alias = CLASSALIASSTRING
End Property
