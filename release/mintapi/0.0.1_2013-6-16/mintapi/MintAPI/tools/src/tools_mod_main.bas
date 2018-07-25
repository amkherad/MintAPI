Attribute VB_Name = "tools_mod_main"
Option Explicit

Public LayerAPI As Object
Public Declare Function InitCommonControlsEx Lib "comctl32.dll" (Iccex As tagInitCommonControlsEx) As Boolean
Public Const ICC_USEREX_CLASSES = &H200

Public Type tagInitCommonControlsEx
   lngSize As Long
   lngICC As Long
End Type

Public Sub Main()
On Error GoTo err
    Dim Iccex As tagInitCommonControlsEx

    With Iccex
           .lngSize = LenB(Iccex)
           .lngICC = ICC_USEREX_CLASSES
    End With

    InitCommonControlsEx Iccex
    
    Set LayerAPI = MintAPI2ndLayerAPI
    
    Call toolsForm.Show
    Exit Sub
err:
    ShowErrorMessage (err.Description)
End Sub

Public Function LanguageEditor() As Object
On Error GoTo err
    Set LanguageEditor = CreateObject("MintAPI2ndLayer.LanguageEditor")
    Exit Function
err:
    ShowErrorMessage (err.Description)
End Function
Public Function MintAPI2ndLayerAPI() As Object
On Error GoTo err
    Set MintAPI2ndLayerAPI = CreateObject("MintAPI2ndLayer.MintAPI2ndLayerAPI")
    Exit Function
err:
    ShowErrorMessage (err.Description)
End Function
Public Function ConfigurationEditor() As Object
On Error GoTo err
    Set ConfigurationEditor = CreateObject("MintAPI2ndLayer.ConfigurationEditor")
    Exit Function
err:
    ShowErrorMessage (err.Description)
End Function
Public Function SourceCodeManager() As Object
On Error GoTo err
    Set SourceCodeManager = CreateObject("MintAPI2ndLayer.SourceCodeManager")
    Exit Function
err:
    ShowErrorMessage (err.Description)
End Function
    
Public Sub ShowErrorMessage(Message As String)
    Call MsgBox("An error occured while trying to connect to MintAPI2ndLayer.dll" & vbCrLf & "Original Error: " & Message, vbCritical, "MintAPI2ndLayer.dll Runtime Error")
End Sub

