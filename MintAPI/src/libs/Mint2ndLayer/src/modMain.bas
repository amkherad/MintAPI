Attribute VB_Name = "modMain"
Option Explicit

Dim i_ModalState As Boolean

Public Sub Main()
    Call Application.EnableVisualStyles(False)
    Call SpecialMethods.RegisterLibraryLicense("", "")
End Sub

Public Sub ShowFormModally(frm As Form, Optional OwnerForm)
    i_ModalState = True
    Call frm.Show(1, OwnerForm)
    i_ModalState = False
End Sub

Public Function ModalState() As Boolean
    ModalState = i_ModalState
End Function
Public Sub CheckNotModalState()
    If i_ModalState Then
        throw Exceptions.Exception("Unable to show a form while a modal forms is showing.")
    End If
End Sub
