VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MintIntellicenseAPI"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'@PROJECT_LICENSE

''<class name="MintIntellicenseAPI" static tags="">
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
Const CLASSID As String = "MintIntellicenseAPI" '#DO NOT REMOVE IT
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
Public Function MenuItemEventArgs(ByVal Sender As Object, ByVal MenuItem As MenuItem) As MenuItemEventArgs
    Set MenuItemEventArgs = New MenuItemEventArgs
    Call MenuItemEventArgs.Constructor0(Sender, MenuItem)
End Function


Public Function MenuItem(ByVal Caption As String, ByVal Parent As MenuItem, Optional ByVal IsPopup As Boolean = False) As MenuItem
    Set MenuItem = mint_main.MenuItem(Caption, Parent, IsPopup)
End Function
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

Public Sub SetVBE(ByVal Value As VBIDE.VBE)
    Set VBInstanceWrapper = New VBInstanceWrapper
    Call VBInstanceWrapper.Constructor0(Value)
End Sub
Public Function VBInstance() As VBInstanceWrapper
    If VBInstanceWrapper Is Nothing Then _
        throw Exps.InvalidStatusException
    Set VBInstance = VBInstanceWrapper
End Function

'*********************************************
'*********************************************
'
'</section>
'---------------------------------------------
'---------------------------------------------
