Attribute VB_Name = "mint_serialization"
Option Explicit

Public Function mint_binaryserialize(ByRef Expression As Variant) As Byte()
    
End Function

Public Function mint_get_byte_array_of(ByRef Expression As Variant, _
        Optional ByVal Length As Long = -1, Optional ByVal Deserializable As Boolean = False) As Byte()
    'BUG BUG BUG BUG BUG BUG
    Dim VT As VariantTypes
    VT = VarType(Expression)
    
    If (VT And vbArray) = vbArray Then '-------------------
        If (VT And vbByte) = vbByte Then
            If Length = -1 Then
                mint_get_byte_array_of = Expression
            Else
                Dim bt() As Byte
                bt = Expression
                mint_get_byte_array_of = GetByteArraySomeLength(bt, Length)
            End If
            Exit Function
        Else
            mint_get_byte_array_of = mint_get_byte_array_of(Expression)
        End If
    ElseIf VT = VBObject Then '-------------------
        Dim itsObject As Object
        Set itsObject = Expression
        If TypeOf itsObject Is ByteArray Then
            Dim BA As ByteArray
            Set BA = itsObject
            mint_get_byte_array_of = BA.ToConstData
        ElseIf TypeOf itsObject Is ObjectBuffer Then
            Dim iobjBuff As ObjectBuffer
            Set iobjBuff = itsObject
        Else
            throw Exps.InvalidArgumentException("at system method mint_get_byte_array_of.")
        End If
    ElseIf VT = vbEmpty Or VT = vbError Or VT = vbDataObject Then
        throw Exps.InvalidArgumentException("at system method mint_get_byte_array_of.")
    ElseIf VT = vbBoolean Then '-------------------
        mint_get_byte_array_of = MemoryToByteArray(CBool(Expression), VLEN_BOOLEAN)
    ElseIf VT = vbByte Then '-------------------
        mint_get_byte_array_of = MemoryToByteArray(CByte(Expression), VLEN_BYTE)
    ElseIf VT = vbCurrency Then '-------------------
        mint_get_byte_array_of = MemoryToByteArray(CCur(Expression), VLEN_CURRENCY)
    ElseIf VT = vbDate Then '-------------------
        mint_get_byte_array_of = MemoryToByteArray(CDate(Expression), VLEN_DATE)
    ElseIf VT = vbDouble Then '-------------------
        mint_get_byte_array_of = MemoryToByteArray(CDbl(Expression), VLEN_DOUBLE)
    ElseIf VT = vbSingle Then '-------------------
        mint_get_byte_array_of = MemoryToByteArray(CSng(Expression), VLEN_SINGLE)
    ElseIf VT = vbInteger Then '-------------------
        mint_get_byte_array_of = MemoryToByteArray(CInt(Expression), VLEN_INTEGER)
    ElseIf VT = vbLong Then '-------------------
        mint_get_byte_array_of = MemoryToByteArray(CLng(Expression), VLEN_LONG)
    ElseIf VT = vbString Then '-------------------
        mint_get_byte_array_of = StringToByteArraySpeed(CStr(Expression))
    ElseIf VT = vbUserDefinedType Then '-------------------
        Dim btArray1() As Byte
        Call CopyMemoryToByteArray(VarPtr(Expression), Len(Expression), btArray1)
        mint_get_byte_array_of = btArray1
    Else '-------------------
        Dim btArray2() As Byte
        If Len(Expression) <= 0 Then throw Exps.InvalidArgumentException("at system method mint_get_byte_array_of.")
        Call CopyMemoryToByteArray(VarPtr(Expression), Len(Expression), btArray2)
        mint_get_byte_array_of = btArray2
    End If

    If Length <> -1 Then
        mint_get_byte_array_of = GetByteArraySomeLength(mint_get_byte_array_of, Length)
    End If
End Function
Public Sub mint_put_byte_array_to(ByRef Target As Variant, ByRef putWhat As Variant, Optional ByVal Length As Long = -1)
    Dim VT As VbVarType
    VT = VarType(Target)
    If (VT And vbArray) = vbArray Then
        If (VT And vbByte) = vbByte Then

        Else

        End If
    ElseIf VT = VBObject Then
        Dim itsObject As Object
        Set itsObject = Target
'        If TypeOf itsObject Is ITexable Then
'            Dim iclsTex As ITexable
'            Set iclsTex = itsObject
'
'        Else
'            throw Exps.UnknownValueException("at system method mint_get_byte_array_of.")
'        End If
    ElseIf VT = vbEmpty Or VT = vbError Then
        throw Exps.InvalidArgumentException("at system method mint_get_byte_array_of.")
    Else

    End If
End Sub
Public Function mint_get_byte_array_of_str(Target, Optional ByVal Length As Long = -1) As Byte()
    Dim VT As VbVarType
    VT = VarType(Target)
    If (VT And vbArray) = vbArray Then '-------------------
        If (VT And vbByte) = vbByte Then
            If Length = -1 Then
                mint_get_byte_array_of_str = Target
            Else
                Dim bt() As Byte
                bt = Target
                mint_get_byte_array_of_str = GetByteArraySomeLength(bt, Length)
            End If
            Exit Function
        Else
            mint_get_byte_array_of_str = mint_binaryserialize(Target)
        End If
    ElseIf VT = VBObject Then '-------------------
        Dim itsObject As Object
        Set itsObject = Target
        If TypeOf itsObject Is ObjectBuffer Then
            Dim iobjBuff As ObjectBuffer
            Set iobjBuff = itsObject
        Else
            throw Exps.InvalidArgumentException("at system method mint_get_byte_array_of_std.")
        End If
    ElseIf VT = vbEmpty Or VT = vbError Or VT = vbDataObject Then
        throw Exps.InvalidArgumentException("at system method mint_get_byte_array_of_std.")
    ElseIf VT = vbBoolean Then '-------------------
        GoTo generic_action
    ElseIf VT = vbByte Then '-------------------
        GoTo generic_action
    ElseIf VT = vbCurrency Then '-------------------
        GoTo generic_action
    ElseIf VT = vbDate Then '-------------------
        GoTo generic_action
    ElseIf VT = vbDouble Then '-------------------
        GoTo generic_action
    ElseIf VT = vbSingle Then '-------------------
        GoTo generic_action
    ElseIf VT = vbInteger Then '-------------------
        GoTo generic_action
    ElseIf VT = vbLong Then '-------------------
        GoTo generic_action
    ElseIf VT = vbString Then '-------------------
        GoTo generic_action
    ElseIf VT = vbUserDefinedType Then '-------------------
        Dim btArray1() As Byte
        Call CopyMemoryToByteArray(VarPtr(Target), Len(Target), btArray1)
        mint_get_byte_array_of_str = btArray1
    Else '-------------------
        Dim btArray2() As Byte
        If Len(Target) <= 0 Then throw Exps.InvalidArgumentException("at system method mint_get_byte_array_of.")
        Call CopyMemoryToByteArray(VarPtr(Target), Len(Target), btArray2)
        mint_get_byte_array_of_str = btArray2
    End If

    If Length <> -1 Then
        mint_get_byte_array_of_str = GetByteArraySomeLength(mint_get_byte_array_of_str, Length)
    End If
    Exit Function
generic_action:
    mint_get_byte_array_of_str = StringToByteArray(CStr(Target))
End Function
Public Sub mint_put_byte_array_to_str(Target, putWhat, Optional ByVal Length As Long = -1)
    Dim VT As VbVarType
    VT = VarType(Target)
    If (VT And vbArray) = vbArray Then
        If (VT And vbByte) = vbByte Then

        Else

        End If
    ElseIf VT = VBObject Then
        Dim itsObject As Object
        Set itsObject = Target
'        If TypeOf itsObject Is ITexable Then
'            Dim iclsTex As ITexable
'            Set iclsTex = itsObject
'
'        Else
'            throw Exps.UnknownValueException("at system method mint_get_byte_array_of_std.")
'        End If
    ElseIf VT = vbEmpty Or VT = vbError Then
        throw Exps.InvalidArgumentException("at system method mint_get_byte_array_of_std.")
    Else

    End If
End Sub
