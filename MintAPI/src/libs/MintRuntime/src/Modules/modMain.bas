Attribute VB_Name = "modMain"
'----------------------------------------------
'  MintAPI by Ali Mousavi Kherad
'  - alimousavikherad@gmail.com
'
'  MintAPI provided under LGPL-v3.
'----------------------------------------------

Option Explicit

Public Const APPLICATIONDOMAIN As String = "com.MintAPI.MintRuntime"

Public Sub Main()
    Dim t As New tApplication
    'Designing and starting you application using Application.StartApplication
    ' will give you some extra features of MintAPI which only availabled
    ' when you set up application to identify you application to MintAPI.
    ' also: you can use Application.InitializeNewApplication to initialize
    '       new application class.
    Call Application.StartApplication(t, DynamicLinkLibrary, App, 0)
End Sub
