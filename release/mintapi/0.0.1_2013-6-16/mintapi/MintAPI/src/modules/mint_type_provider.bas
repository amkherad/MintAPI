Attribute VB_Name = "mint_type_provider"
Option Explicit

Private Type IPROVIDERS
    ProviderEngine As String
    Alias As String
    Provider As IProvider
End Type

Dim prvs() As IPROVIDERS
Dim prvsCount As Long

Private Function ValidateProviderEngine(ProviderEngine As String) As Boolean
    
End Function
Private Function ValidateAlias(Alias As String) As Boolean
    
End Function

Public Sub RegisterProvider(ProviderEngine As String, Alias As String, Provider As IProvider)
    
End Sub
Public Sub ChangeProvider(ProviderEngine As String, Alias As String, Provider As IProvider)
    
End Sub
Public Sub UnRegisterProvider(ProviderEngine As String, Alias As String)
    
End Sub

Public Function Provide(ProviderEngine As String, Alias As String, Args As ArgumentList, ByRef retVal) As Boolean
    
End Function
