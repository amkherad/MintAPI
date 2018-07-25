Attribute VB_Name = "bitOperations"
'@PROJECT_LICENSE

'by Ali Mousavi Kherad

Option Explicit
Option Base 0
Const CLASSID As String = "bitOperations"

Dim inited As Boolean
Dim bitMask(0 To 31) As Long

Public Sub Initialize()
    If inited Then Exit Sub
    Dim i As Long
    For i = 0 To 30
        bitMask(i) = 2 ^ i
    Next 'creating event mask
    bitMask(31) = &H80000000
    inited = True
End Sub
Public Sub Dispose(Optional ByVal Force As Boolean = False)
    If Not inited Then Exit Sub
    Erase bitMask
    inited = False
End Sub

Public Function Power2(ByVal Exponent As Long) As Long
    If Not inited Then Call Initialize
    ' rule out errors
    If Exponent < 0 Or Exponent > 31 Then throw Exps.IndexOutOfRangeException
    ' initialize the array at the first call
    ' return the result
    Power2 = bitMask(Exponent)
End Function

' Rotate a Long to the left the specified number of times
Public Function RotateLeft(ByVal Value As Long, ByVal Times As Long) As Long
    Dim i As Long, SignBits As Long
    ' no need to rotate more times than required
    Times = Times Mod 32
    ' return the number if it's a multiple of 32
    If Times = 0 Then RotateLeft = Value: Exit Function
    For i = 1 To Times
        ' remember the 2 most significant bits
        SignBits = Value And &HC0000000
        ' clear those bit and shift to the left by one position
        Value = (Value And &H3FFFFFFF) * 2
        ' if the number was negative, then add 1
        ' if bit 30 was set, then set the sign bit
        Value = Value Or ((SignBits < 0) And 1) Or (CBool(SignBits And &H40000000) And &H80000000)
    Next
    RotateLeft = Value
End Function

' Rotate an Integer to the left the specified number of times
Public Function RotateLeftI(ByVal Value As Integer, ByVal Times As Long) As Integer
    Dim i As Long, SignBits As Integer

    ' no need to rotate more times than required
    Times = Times Mod 16
    ' return the number if it's a multiple of 16
    If Times = 0 Then RotateLeftI = Value: Exit Function

    For i = 1 To Times
        ' remember the 2 most significant bits
        SignBits = Value And &HC000
        ' clear those bit and shift to the left by one position
        Value = (Value And &H3FFF) * 2
        ' if the number was negative, then add 1
        ' if bit 30 was set, then set the sign bit
        Value = Value Or ((SignBits < 0) And 1) Or (CBool(SignBits And &H4000) And &H8000)
    Next
    RotateLeftI = Value
End Function

' Rotate a Long to the right the specified number of times
Public Function RotateRight(ByVal Value As Long, ByVal Times As Long) As Long
    Dim i As Long, SignBits As Long

    ' no need to rotate more times than required
    Times = Times Mod 32
    ' return the number if it's a multiple of 32
    If Times = 0 Then RotateRight = Value: Exit Function

    For i = 1 To Times
        ' remember the sign bit and bit 0
        SignBits = Value And &H80000001
        ' clear those bits and shift to the right by one position
        Value = (Value And &H7FFFFFFE) \ 2
        ' if the number was negative, then re-insert the bit
        ' if bit 0 was set, then set the sign bit
        Value = Value Or ((SignBits < 0) And &H40000000) Or (CBool(SignBits And 1) And &H80000000)
    Next
    RotateRight = Value
End Function

' Rotate an Integer to the right the specified number of times
Public Function RotateRightI(ByVal Value As Integer, ByVal Times As Long) As Integer
    Dim i As Long, SignBits As Integer

    ' no need to rotate more times than required
    Times = Times Mod 16
    ' return the number if it's a multiple of 16
    If Times = 0 Then RotateRightI = Value: Exit Function

    For i = 1 To Times
        ' remember the sign bit and bit 0
        SignBits = Value And &H8001
        ' clear those bits and shift to the right by one position
        Value = (Value And &H7FFE) \ 2
        ' if the number was negative, then re-insert the bit
        ' if bit 0 was set, then set the sign bit
        Value = Value Or ((SignBits < 0) And &H4000) Or (CBool(SignBits And 1) And &H8000)
    Next
    RotateRightI = Value
End Function

' Shift to the left of the specified number of times
Public Function ShiftLeft(ByVal Value As Long, ByVal Times As Long) As Long
    ' we need to create a mask of 1's corresponding to the
    ' times in VALUE that will be retained in the result
    Dim Mask As Long, SignBit As Long

    ' return zero if too many times
    If Times >= 32 Then Exit Function
    ' return the value if zero times
    If Times = 0 Then ShiftLeft = Value: Exit Function

    ' this extracts the bit in Value that will become the sign bit
    Mask = Power2(31 - Times)
    ' this calculates the sign bit of the result
    SignBit = CBool(Value And Mask) And &H80000000
    ' this clears all the most significant times,
    ' that would be lost anyway, and also clears the sign bit
    Value = Value And (Mask - 1)
    ' do the shift to the left, without risking an overflow
    ' and then add the sign bit
    ShiftLeft = (Value * Power2(Times)) Or SignBit
End Function

' Shift to the right of the specified number of times
Public Function ShiftRight(ByVal Value As Long, ByVal Times As Long) As Long
    ' we need to create a mask of 1's corresponding to the
    ' digits in VALUE that will be retained in the result
    Dim Mask As Long, SignBit As Long
    ' return zero if too many times
    If Times >= 32 Then Exit Function
    ' return the value if zero times
    If Times = 0 Then ShiftRight = Value: Exit Function
    ' evaluate the sign bit in advance
    SignBit = (Value < 0) And Power2(31 - Times)
    ' create a mask with 1's for the digits that will be preserved
    If Times < 31 Then
        ' if times>=31 then the mask is zero
        Mask = Not (Power2(Times) - 1)
    Else
        Exit Function
    End If
    ' clear all the digits that will be discarded, and
    ' also clear the sign bit
    Value = (Value And &H7FFFFFFF) And Mask
    ' do the shift, without any problem, and add the sign bit
    ShiftRight = (Value \ Power2(Times)) Or SignBit
End Function
