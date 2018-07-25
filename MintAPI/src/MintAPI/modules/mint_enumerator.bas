Attribute VB_Name = "mint_enumerator"
Option Explicit

Private Const ENUM_FINISHED As Long = 1

Private IID_IUnknown As StdGuid
Private Const IID_IUnknown_Data1 As Long = 0
Private IID_IEnumVariant As StdGuid
Private Const IID_IEnumVariant_Data1 As Long = &H20404

Private Type UserEnumeratorWrapperType
    pVTable As Long
    cRefs As Long
    UserEnum As IEnumerator
End Type
Private Type VTable
    Functions(0 To 6) As Long
End Type

Private mVTable As VTable
Private mpVTable As Long

Private Sub InitializeEnumeratorSubSystem()
    Call InitializeGUIDS
    Call InitializeVirtualTable
End Sub

Private Sub InitializeGUIDS()
    With IID_IEnumVariant
        .Data1 = &H20404
        .Data4(0) = &HC0
        .Data4(7) = &H46
    End With
    With IID_IUnknown
        .Data4(0) = &HC0
        .Data4(7) = &H46
    End With
End Sub
Private Sub InitializeVirtualTable()
    With mVTable
        .Functions(0) = GetAddressOf(AddressOf QueryInterface)
        .Functions(1) = GetAddressOf(AddressOf AddRef)
        .Functions(2) = GetAddressOf(AddressOf Release)
        .Functions(3) = GetAddressOf(AddressOf IEnumVariant_Next)
        .Functions(4) = GetAddressOf(AddressOf IEnumVariant_Skip)
        .Functions(5) = GetAddressOf(AddressOf IEnumVariant_Reset)
        .Functions(6) = GetAddressOf(AddressOf IEnumVariant_Clone)
        
        mpVTable = VarPtr(.Functions(0))
   End With
End Sub

Public Function CreateEnumerator(ByVal ObjEnumerator As IEnumerator) As IUnknown
    Dim This As Long
    Dim Struct As UserEnumeratorWrapperType
    
    If mpVTable = 0 Then Call InitializeEnumeratorSubSystem
    
    ' allocate memory to place the new object.
    This = API_CoTaskMemAlloc(Len(Struct))
    If This = vbNullPtr Then throw Exps.OutOfMemoryException
    
    ' fill the structure of the new wrapper object
    With Struct
        Set .UserEnum = ObjEnumerator
        .cRefs = 1
        .pVTable = mpVTable
    End With
    
    ' move the structure to the allocated memory to complete the object
    Call memcpy(ByVal This, ByVal VarPtr(Struct), Len(Struct))
    Call memzero(ByVal VarPtr(Struct), Len(Struct))
    
    ' assign the return value to the newly create object.
    ObjectPtr(CreateEnumerator) = This
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''  VTable functions in the IEnumVariant and IUnknown interfaces.            ''
''  by Kelly Ethridge (VBCorLib) - 2004 GNU Library General Public License   ''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''  I knew these myself :-) '
'''''''''''''''''''''''''''''

' When VB queries the interface, we support only two.
' IUnknown
' IEnumVariant
Private Function QueryInterface(ByRef This As UserEnumeratorWrapperType, _
                                ByRef riid As StdGuid, _
                                ByRef pvObj As Long) As Long
    Dim OK As Long
    
    Select Case riid.Data1
        Case IID_IEnumVariant_Data1
            OK = API_IsEqualGUID(riid, IID_IEnumVariant)
        Case IID_IUnknown_Data1
            OK = API_IsEqualGUID(riid, IID_IUnknown)
    End Select
    
    If OK Then
        pvObj = VarPtr(This)
        Call AddRef(This)
    Else
        QueryInterface = E_NOINTERFACE
    End If
End Function


' increment the number of references to the object.
Private Function AddRef(ByRef This As UserEnumeratorWrapperType) As Long
    With This
        Dim cRefs As Long
        cRefs = .cRefs + 1
        .cRefs = cRefs
        AddRef = cRefs
    End With
End Function


' decrement the number of references to the object, checking
' to see if the last reference was released.
Private Function Release(ByRef This As UserEnumeratorWrapperType) As Long
    With This
        Dim cRefs As Long
        cRefs = .cRefs - 1
        .cRefs = cRefs
        Release = cRefs
        If cRefs = 0 Then Call Delete(This)
    End With
End Function


' cleans up the lightweight objects and releases the memory
Private Sub Delete(ByRef This As UserEnumeratorWrapperType)
   Set This.UserEnum = Nothing
   Call API_CoTaskMemFree(VarPtr(This))
End Sub


' move to the next element and return it, signaling if we have reached the end.
Private Function IEnumVariant_Next(ByRef This As UserEnumeratorWrapperType, ByVal celt As Long, ByRef prgVar As Variant, ByVal pceltFetched As Long) As Long
    If This.UserEnum.MoveNext Then
        Call MoveVariant(prgVar, This.UserEnum.Current)
        
        ' check to see if the pointer is valid (not zero)
        ' before we write to that memory location.
        If pceltFetched Then _
            MemLong(pceltFetched) = 1
    Else
        IEnumVariant_Next = ENUM_FINISHED
    End If
End Function


' skip the requested number of elements as long as we don't run out of them.
Private Function IEnumVariant_Skip(ByRef This As UserEnumeratorWrapperType, ByVal celt As Long) As Long
    Do While celt > 0
        If This.UserEnum.MoveNext = False Then
            IEnumVariant_Skip = ENUM_FINISHED
            Exit Function
        End If
        celt = celt - 1
    Loop
End Function


' request the user enum to reset.
Private Function IEnumVariant_Reset(ByRef This As UserEnumeratorWrapperType) As Long
   Call This.UserEnum.Reset
End Function


' we just return a reference to the original object.
Private Function IEnumVariant_Clone(ByRef This As UserEnumeratorWrapperType, ByRef ppEnum As IUnknown) As Long
    Dim O As ICloneable
    Set O = This.UserEnum
    Set ppEnum = O.Clone
End Function

