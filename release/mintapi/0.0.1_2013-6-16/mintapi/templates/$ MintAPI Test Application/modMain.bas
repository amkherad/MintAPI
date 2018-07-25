Attribute VB_Name = "modMain"
'----------------------------------------------
'  MintAPI by Ali Mousavi Kherad
'  - alimousavikherad@gmail.com
'
'  MintAPI provided under LGPL-v3.
'----------------------------------------------

Option Explicit

Public Sub Main()
    'Only stdout refers to console!
    Call Application.trigConsole
    
    'Used to enter application in developermode ,this may come usefull
    ' when you are at design time.
    'Call Application.EnterDeveloperMode
    
    Dim t As New tApplication
    'Designing and starting you application using Application.StartApplication
    ' will give you some extra features of MintAPI which only availabled
    ' when you set up application to identify you application to MintAPI.
    ' also: you can use Application.InitializeNewApplication to initialize
    '       new application class.
    Call Application.StartApplication(t, ApplicationExecutable, App, 0)
End Sub
