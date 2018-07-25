VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ProviderStatic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'@PROJECT_LICENSE

''<class name="ProviderStatic" static tags="">
''  <summary>
''    Some class description.
''  </summary>
''  <using>Some class alias.</using>
''  <idea>Me</idea>
''  <sample></sample>
''</class>

Option Base 0
Option Explicit

'=============================================
'=============================================
'=============================================
'<section Meta Data>
'
Const CLASSID As String = "ProviderStatic" '#DO NOT REMOVE IT
Const CLASSALIAS As String = APPLICATIONDOMAIN & "." & CLASSID '#DO NOT REMOVE IT
Const CLASSVERSIONSTRING As String = "0.1" '#DO NOT REMOVE IT
Const CLASSALIASSTRING As String = CLASSALIAS & "/" & CLASSVERSIONSTRING '#DO NOT REMOVE IT

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
'<section Creators>
'
Public Property Get DefaultGlobalizationProvider() As Object 'DynamicObject
    Dim GP As IProvider
    
    Set DefaultGlobalizationProvider = DynamicObject.Create(GP)
End Property
Public Property Get DefaultDataProvider() As Object 'DynamicObject
    Dim DP As IProvider
    
    Set DefaultDataProvider = DynamicObject.Create(DP)
End Property
Public Property Get DefaultResourceProvider() As Object 'DynamicObject
    Dim RP As IProvider
    
    Set DefaultResourceProvider = DynamicObject.Create(RP)
End Property
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

Public Sub RegisterProvider(ByVal ProviderEngine As String, ByVal Alias As String, ByVal Provider As IProvider)
    
End Sub
Public Sub UnRegisterProvider(ByVal ProviderEngine As String, ByVal Alias As String)
    
End Sub
Public Function Exists(ByVal ProviderEngine As String, ByVal Alias As String) As Boolean
    
End Function

Public Function Provide(ByVal ProviderEngine As String, ByVal Alias As String, ByVal Args As ArgumentList, outRetVal As Variant) As Boolean
    Select Case LCase(ProviderEngine)
        Case vbNullString, "cp", "classprovider", "class_provider"
            Set outRetVal = ProvideClass(Alias, Args)
            Provide = True
        Case "mint_service_provider"
            
        Case Else
            throw Exps.InvalidArgumentException("Unknown provider engine.")
    End Select
End Function

Friend Function ProvideClass(ByVal Alias As String, ByVal Args As ArgumentList) As Object
    'Alias : [[com.]mintapi.]ByteArray[/1.0]
    
End Function

'*********************************************
'*********************************************
'
'</section>
'---------------------------------------------
'---------------------------------------------
