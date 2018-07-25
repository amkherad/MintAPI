Attribute VB_Name = "mint_main"
Option Explicit

Public Const APPLICATIONDOMAIN As String = "MintAPIIntellicense"

Public Sub Main()
    
End Sub

Public Function MenuItemEventArgs(ByVal Sender As Object, ByVal MenuItem As MenuItem) As MenuItemEventArgs
    Set MenuItemEventArgs = New MenuItemEventArgs
    Call MenuItemEventArgs.Constructor0(Sender, MenuItem)
End Function

Public Function MenuItem(ByVal Caption As String, ByVal Parent As MenuItem, Optional ByVal IsPopup As Boolean = False) As MenuItem
    Set MenuItem = New MenuItem
    MenuItem.Caption = Caption
    MenuItem.IsParent = IsPopup
    Call Parent.Children.Add(MenuItem)
End Function
