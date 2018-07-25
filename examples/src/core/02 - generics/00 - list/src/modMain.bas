Attribute VB_Name = "modMain"
'----------------------------------------------
'  MintAPI by Ali Mousavi Kherad
'  - alimousavikherad@gmail.com
'
'  MintAPI provided under LGPL-v3.
'----------------------------------------------

Option Explicit
Public Declare Function VarPtrArray Lib "msvbvm60" Alias "VarPtr" (Ptr() As Any) As Long
Public Declare Sub VariantCopy Lib "oleaut32" (ByRef pvargDest As Variant, ByRef pvargSrc As Variant)
Public Declare Sub VariantClear Lib "oleaut32" (ByRef pvarg As Variant)

Public Sub Main(): InitializeEnvironment (Command$)
    Dim t As New tApplication
    'Designing and starting you application using Application.Run()
    ' will give you some extra features of MintAPI which only available
    ' when you set up an application class to identify you app to MintAPI.
    ' also: you can use Application.InitializeNewApplication to initialize
    '       new application class.
    Call Application.Run(t, ApplicationExecutable Or OwnConsole, App, 0)
End Sub

Public Sub InitializeEnvironment(ByVal CommandLine As String)
    'Put your basic initializations here...
End Sub
