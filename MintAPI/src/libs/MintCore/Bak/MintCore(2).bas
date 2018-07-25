'freeBasic

extern "C"
    public sub InitializeMintCore() export
        
    end sub
    public sub DisposeMintCore() export
        
    end sub
    public sub initialize_wxWidgets(dllPath as zstring) export
        
    end sub
end extern

extern "C"
    public function DllRegisterServer() as long export
        return 0
    end function
    public function DllUnRegisterServer() as long export
        return 0
    end Function
end extern