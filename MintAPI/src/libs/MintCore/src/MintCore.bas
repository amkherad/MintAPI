'freeBasic

Extern "C"
    Public Sub InitializeMintCore() Export
        
    End Sub
    Public Sub DisposeMintCore() Export
        
    End Sub
    Public Sub initialize_wxWidgets(dllPath As ZString) Export
        
    End Sub
End Extern

Extern "C"
    Public Function DllRegisterServer() As Long Export
        return 0
    End function
    Public Function DllUnRegisterServer() As Long Export
        return 0
    End Function
End Extern