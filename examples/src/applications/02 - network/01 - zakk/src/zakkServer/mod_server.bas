Attribute VB_Name = "mod_server"
Option Explicit

Public pServer As Server

Public Sub InitializeServer(Command As String)
    If pServer Is Nothing Then throw Exceptions.Exception("Server already initialized.")
    Set pServer = New Server
    Call pServer.Initialize(Command)
End Sub

