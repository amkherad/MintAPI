'freeBasic

Extern "C"
    Public Sub InitializeMintCore() export
        
    End Sub
    Public Sub DisposeMintCore() export
        
    End Sub
    Public Sub initialize_wxWidgets(dllPath as zstring) export
        
    End Sub
End Extern

Extern "C"
    Public Function DllRegisterServer() as long export
        return 0
    End function
    Public Function DllUnRegisterServer() as long export
        return 0
    End Function
End Extern