Extern "C"
    Public Sub memcpy1(ByVal Length As Integer, ByVal Source As Any Ptr , ByVal Destination As  Any Ptr) Export
        CopyMemory(Destination, Source, Length)
    End Sub
End Extern