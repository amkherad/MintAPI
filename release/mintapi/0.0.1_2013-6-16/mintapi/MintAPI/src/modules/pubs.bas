Attribute VB_Name = "pubs"
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Copyright (C) 2013 Ali Mousavi Kherad and/or other contributors.
'' Contact: alimousavikherad@gmail.com http://www.sourceforge.net/users/amgmail
''
'' This file is part of the MintAPI dll of the MintAPI Toolkit.
''
'' $_BEGIN_LICENSE:LGPL$
''
'' GNU Lesser General Public License Usage
'' Alternatively, this file may be used under the terms of the GNU Lesser
'' General Public License version 2.1 as published by the Free Software
'' Foundation and appearing in the file LICENSE.LGPL included in the
'' packaging of this file.  Please review the following information to
'' ensure the GNU Lesser General Public License version 2.1 requirements
'' will be met: http://www.gnu.org/licenses/old-licenses/lgpl-2.1.html.
''
'' In addition, as a special exception, MintAPI gives you certain additional
'' rights.  These rights are described in the MintAPI LGPL Exception
'' version 1.1, included in the file LGPL_EXCEPTION.txt in this package.
''
'' GNU General Public License Usage
'' Alternatively, this file may be used under the terms of the GNU
'' General Public License version 3.0 as published by the Free Software
'' Foundation and appearing in the file LICENSE.GPL included in the
'' packaging of this file.  Please review the following information to
'' ensure the GNU General Public License version 3.0 requirements will be
'' met: http://www.gnu.org/copyleft/gpl.html.
''
'' $_END_LICENSE$
''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

'by Ali Mousavi Kherad (LGPL-v3)
'Free to use and distribute but including my name as Ali Mousavi Kherad and email as (alimousavikherad@gmail.com)!
Public mint_api_dialogs_last_choose_file_read_only_flag_state As Boolean
'Public mint_api_console_is_breaked As Boolean
Public mint_api_console_instances As Long
Public mint_api_console_out_handle As Long
Public mint_api_console_in_handle As Long
Public mint_api_console_err_handle As Long

Public Const OBJECTBUFFER_STREAMINGVALUE As String = "streamingvalue"
Public Const OBJECTBUFFER_HANDLE As String = "handle"
Public Const OBJECTBUFFER_DISPLAY_CONTEXT As String = "hdc"
Public Const OBJECTBUFFER_EVENTNAMES_LIST As String = "eventnameslist"
Public Const OBJECTBUFFER_TEXEDOBJECT As String = "itexedobject"
Public Const OBJECTBUFFER_MINTLOCALTYPE As String = "mintlocaltype"
Public Const OBJECTBUFFER_INHERIT As String = "inherit"
Public Const OBJECTBUFFER_RENDERTARGET As String = "rendertarget"
Public Const OBJECTBUFFER_SIGNAL As String = "signal"
Public Const OBJECTBUFFER_SLOT As String = "slot"
Public Const OBJECTBUFFER_NULL As String = "null"
Public Const OBJECTBUFFER_CONTROLLERINSTANCE As String = "controllerinstance"

Public Const OBJECTBUFFER_VALUES_STREAMING_CLEAR As String = "clear"
Public Const OBJECTBUFFER_VALUES_STREAMING_NEWLINE As String = "newline"
Public Const OBJECTBUFFER_VALUES_STREAMING_FLUSH As String = "flush"
Public Const OBJECTBUFFER_VALUES_STREAMING_SEEK As String = "seek"

Public Type EventArgs_InnerData
    ArgsCount As Long
    Args() As Variant
End Type

Public Function IsEmptyArray(targetArray) As Boolean
If Not (VarType(targetArray) And vbArray) = vbArray Then throw InvalidArgumentTypeException
    On Error GoTo zeroLength
    IsEmptyArray = (UBound(targetArray) - LBound(targetArray) + 1) <= 0
    Exit Function
zeroLength:
    IsEmptyArray = True
End Function
Public Sub EmptyVar(Variable)
    If IsObject(Variable) Then
        Set Variable = Nothing
    Else
        If IsArray(Variable) Then
            Erase Variable
        Else
            Variable = Empty
        End If
    End If
End Sub
Public Function IsEmptyVariable(targetVariable) As Boolean
    Select Case VarType(targetVariable)
        Case vbArray
            IsEmptyVariable = (IsArrayEmpty(targetVariable))
        Case VBObject
            IsEmptyVariable = (targetVariable Is Nothing)
        Case Else
            IsEmptyVariable = (targetVariable = Empty)
    End Select
End Function

Public Function SelectArg(StaticVar, Optional OptionalVar) As Variant
    If IsMissing(OptionalVar) Then
        If IsObject(StaticVar) Then
            Set SelectArg = StaticVar
        Else
                SelectArg = StaticVar
        End If
    Else
        If IsObject(OptionalVar) Then
            Set SelectArg = OptionalVar
        Else
                SelectArg = OptionalVar
        End If
    End If
End Function
Public Function SelectString(StaticStr As String, Optional OptionalStr) As String
    If IsMissing(OptionalStr) Then
        SelectString = StaticStr
    Else
        If IsObject(OptionalStr) Then _
            throw InvalidArgumentTypeException("Invalid OptionalStr type.")
        SelectString = CStr(OptionalStr)
    End If
End Function


Public Function Select_Arg(First As String, Second As String, Optional ByVal ChooseFirst As Boolean = True) As String
    If ChooseFirst Then
        Select_Arg = First
    Else
        Select_Arg = Second
    End If
End Function



Public Function GlobalFilters(Optional includes, Optional excludes) As GlobalFilters
    Dim retVal As GlobalFilters
    If IsMissing(includes) Then
        Dim p(1) As String
        p(0) = "*": p(1) = "*.*"
        includes = p
    End If
    If VarType(includes) = (vbArray Or vbString) Then
        retVal.IncludeTemplates = includes
    Else
        ReDim retVal.IncludeTemplates(0)
        retVal.IncludeTemplates(0) = CStr(includes)
    End If
    If Not IsMissing(excludes) Then
        If VarType(excludes) = (vbArray Or vbString) Then
            retVal.ExcludeTemplates = excludes
        Else
            ReDim retVal.ExcludeTemplates(0)
            retVal.ExcludeTemplates(0) = CStr(excludes)
        End If
    End If
    GlobalFilters = retVal
End Function
Public Function FileFilters(Optional includes, Optional excludes) As GlobalFilters
    Dim retVal As GlobalFilters
    If IsMissing(includes) Then
        Dim p(1) As String
        p(0) = "*": p(1) = "*.*"
        includes = p
    End If
    If VarType(includes) = (vbArray Or vbString) Then
        retVal.IncludeTemplates = includes
    Else
        ReDim retVal.IncludeTemplates(0)
        retVal.IncludeTemplates(0) = CStr(includes)
    End If
    If Not IsMissing(excludes) Then
        If VarType(excludes) = (vbArray Or vbString) Then
            retVal.ExcludeTemplates = excludes
        Else
            ReDim retVal.ExcludeTemplates(0)
            retVal.ExcludeTemplates(0) = CStr(excludes)
        End If
    End If
    FileFilters = retVal
End Function
Public Function StringArray(ParamArray str() As Variant) As String()
    Dim i As Long, strSize As Long, retVal() As String
    On Error GoTo zeroLength
    strSize = UBound(str) - LBound(str) + 1
    If strSize > 0 Then
        ReDim retVal(strSize - 1)
        For i = 0 To strSize - 1
            retVal(i) = CStr(str(i))
        Next
    End If
    StringArray = retVal
zeroLength:
End Function
Public Function IncludeFilter_single(fs() As String, Expression As String) As Boolean
    Dim i As Long
    For i = 0 To ArraySize(fs) - 1
        If (Expression) Like CStr(fs(i)) Then
            IncludeFilter_single = True
            Exit Function
        End If
    Next
    IncludeFilter_single = False
End Function
Public Function IsFilterIncluded(Filter As GlobalFilters, Expression As String) As Boolean
    IsFilterIncluded = ((IncludeFilter_single(Filter.IncludeTemplates, Expression)) And (Not IncludeFilter_single(Filter.ExcludeTemplates, Expression)))
End Function


Public Function InstallPlugin(Path As String) As Long

End Function
Public Function LoadPlugin(Path As String) As IPlugin

End Function
Public Function LoadPluginFromFile(Path As String) As IPlugin

End Function

Private Function RegUnregActiveX(Path As String, Optional Register As Boolean = True) As Long

End Function
Public Function RegisterActiveX(Path As String) As Long
    RegisterActiveX = RegUnregActiveX(Path)
End Function
Public Function UnRegisterActiveX(Path As String) As Long
    UnRegisterActiveX = RegUnregActiveX(Path, False)
End Function




Public Function mint_get_byte_array_of(target, Optional ByVal Length As Long = -1) As Byte()
    Dim vt As VbVarType
    vt = VarType(target)
    If (vt And vbArray) = vbArray Then '-------------------
        If (vt And vbByte) = vbByte Then
            If Length = -1 Then
                mint_get_byte_array_of = target
            Else
                Dim bt() As Byte
                bt = target
                mint_get_byte_array_of = GetByteArraySomeLength(bt, Length)
            End If
            Exit Function
        Else
            mint_get_byte_array_of = ArrayToByteArray(target)
        End If
    ElseIf vt = VBObject Then '-------------------
        Dim itsObject As Object
        Set itsObject = target
        If TypeOf itsObject Is ByteArray Then
            Dim ba As ByteArray
            Set ba = itsObject
            mint_get_byte_array_of = ba.constData
        ElseIf TypeOf itsObject Is IClassTexer Then
            Dim iclsTex As IClassTexer
            Set iclsTex = itsObject
            mint_get_byte_array_of = iclsTex.toByteArray
        ElseIf TypeOf itsObject Is ObjectBuffer Then
            Dim iobjBuff As ObjectBuffer
            Set iobjBuff = itsObject
        Else
            throw UnknownValueException("at system method mint_get_byte_array_of.")
        End If
    ElseIf vt = vbEmpty Or vt = vbError Or vt = vbDataObject Then
        throw InvalidArgumentTypeException("at system method mint_get_byte_array_of.")
    ElseIf vt = vbBoolean Then '-------------------
        Dim vtBoolean As Boolean
        vtBoolean = target
        mint_get_byte_array_of = BooleanToByteArray(vtBoolean)
    ElseIf vt = vbByte Then '-------------------
        Dim vtByte As Byte
        vtByte = target
        mint_get_byte_array_of = ByteToByteArray(vtByte)
    ElseIf vt = vbCurrency Then '-------------------
        Dim vtCurrency As Currency
        vtCurrency = target
        mint_get_byte_array_of = CurrencyToByteArray(vtCurrency)
    ElseIf vt = vbDate Then '-------------------
        Dim vtDate As Date
        vtDate = target
        mint_get_byte_array_of = DateToByteArray(vtDate)
    ElseIf vt = vbDouble Then '-------------------
        Dim vtDouble As Double
        vtDouble = target
        mint_get_byte_array_of = DoubleToByteArray(vtDouble)
    ElseIf vt = vbSingle Then '-------------------
        Dim vtSingle As Single
        vtSingle = target
        mint_get_byte_array_of = SingleToByteArray(vtSingle)
    ElseIf vt = vbInteger Then '-------------------
        Dim vtInteger As Integer
        vtInteger = target
        mint_get_byte_array_of = IntegerToByteArray(vtInteger)
    ElseIf vt = vbLong Then '-------------------
        Dim vtLong As Long
        vtLong = target
        mint_get_byte_array_of = LongToByteArray(vtLong)
    ElseIf vt = vbString Then '-------------------
        Dim vtString As String
        vtString = target
        mint_get_byte_array_of = StringToByteArray(vtString)
    ElseIf vt = vbUserDefinedType Then '-------------------
        Dim btArray1() As Byte
        Call CopyMemoryToByteArray(VarPtr(target), Len(target), btArray1)
        mint_get_byte_array_of = btArray1
    Else '-------------------
        Dim btArray2() As Byte
        If Len(target) <= 0 Then throw InvalidArgumentTypeException("at system method mint_get_byte_array_of.")
        Call CopyMemoryToByteArray(VarPtr(target), Len(target), btArray2)
        mint_get_byte_array_of = btArray2
    End If

    If Length <> -1 Then
        mint_get_byte_array_of = GetByteArraySomeLength(mint_get_byte_array_of, Length)
    End If
End Function
Public Sub mint_put_byte_array_to(target, putWhat, Optional ByVal Length As Long = -1)
    Dim vt As VbVarType
    vt = VarType(target)
    If (vt And vbArray) = vbArray Then
        If (vt And vbByte) = vbByte Then

        Else

        End If
    ElseIf vt = VBObject Then
        Dim itsObject As Object
        Set itsObject = target
        If TypeOf itsObject Is IClassTexer Then
            Dim iclsTex As IClassTexer
            Set iclsTex = itsObject

        Else
            throw UnknownValueException("at system method mint_get_byte_array_of.")
        End If
    ElseIf vt = vbEmpty Or vt = vbError Then
        throw InvalidArgumentTypeException("at system method mint_get_byte_array_of.")
    Else

    End If
End Sub
Public Function mint_get_byte_array_of_std(target, Optional ByVal Length As Long = -1) As Byte()
    Dim vt As VbVarType
    vt = VarType(target)
    If (vt And vbArray) = vbArray Then '-------------------
        If (vt And vbByte) = vbByte Then
            If Length = -1 Then
                mint_get_byte_array_of_std = target
            Else
                Dim bt() As Byte
                bt = target
                mint_get_byte_array_of_std = GetByteArraySomeLength(bt, Length)
            End If
            Exit Function
        Else
            mint_get_byte_array_of_std = ArrayToByteArray(target)
        End If
    ElseIf vt = VBObject Then '-------------------
        Dim itsObject As Object
        Set itsObject = target
        If TypeOf itsObject Is IClassTexer Then
            Dim iclsTex As IClassTexer
            Set iclsTex = itsObject
            mint_get_byte_array_of_std = iclsTex.toByteArray
        ElseIf TypeOf itsObject Is ObjectBuffer Then
            Dim iobjBuff As ObjectBuffer
            Set iobjBuff = itsObject
        Else
            throw UnknownValueException("at system method mint_get_byte_array_of_std.")
        End If
    ElseIf vt = vbEmpty Or vt = vbError Or vt = vbDataObject Then
        throw InvalidArgumentTypeException("at system method mint_get_byte_array_of_std.")
    ElseIf vt = vbBoolean Then '-------------------
        GoTo generic_action
    ElseIf vt = vbByte Then '-------------------
        GoTo generic_action
    ElseIf vt = vbCurrency Then '-------------------
        GoTo generic_action
    ElseIf vt = vbDate Then '-------------------
        GoTo generic_action
    ElseIf vt = vbDouble Then '-------------------
        GoTo generic_action
    ElseIf vt = vbSingle Then '-------------------
        GoTo generic_action
    ElseIf vt = vbInteger Then '-------------------
        GoTo generic_action
    ElseIf vt = vbLong Then '-------------------
        GoTo generic_action
    ElseIf vt = vbString Then '-------------------
        GoTo generic_action
    ElseIf vt = vbUserDefinedType Then '-------------------
        Dim btArray1() As Byte
        Call CopyMemoryToByteArray(VarPtr(target), Len(target), btArray1)
        mint_get_byte_array_of_std = btArray1
    Else '-------------------
        Dim btArray2() As Byte
        If Len(target) <= 0 Then throw InvalidArgumentTypeException("at system method mint_get_byte_array_of.")
        Call CopyMemoryToByteArray(VarPtr(target), Len(target), btArray2)
        mint_get_byte_array_of_std = btArray2
    End If

    If Length <> -1 Then
        mint_get_byte_array_of_std = GetByteArraySomeLength(mint_get_byte_array_of_std, Length)
    End If
    Exit Function
generic_action:
    mint_get_byte_array_of_std = StringToByteArray(CStr(target))
End Function
Public Sub mint_put_byte_array_to_std(target, putWhat, Optional ByVal Length As Long = -1)
    Dim vt As VbVarType
    vt = VarType(target)
    If (vt And vbArray) = vbArray Then
        If (vt And vbByte) = vbByte Then

        Else

        End If
    ElseIf vt = VBObject Then
        Dim itsObject As Object
        Set itsObject = target
        If TypeOf itsObject Is IClassTexer Then
            Dim iclsTex As IClassTexer
            Set iclsTex = itsObject

        Else
            throw UnknownValueException("at system method mint_get_byte_array_of_std.")
        End If
    ElseIf vt = vbEmpty Or vt = vbError Then
        throw InvalidArgumentTypeException("at system method mint_get_byte_array_of_std.")
    Else

    End If
End Sub


Public Sub mint_setstream_state(State As Boolean, inoutState As Boolean, State_LOCK As String, Optional Reserved)
    Dim strReserved As String
    strReserved = IIf(IsMissing(Reserved), "", CStr(Reserved))
    If State Then
        If inoutState Then
            If State_LOCK <> strReserved Then GoTo errThrow
        Else
            If State_LOCK <> "" Then
                If strReserved <> State_LOCK Then GoTo errThrow
            End If
            inoutState = True
            State_LOCK = strReserved
        End If
    Else
        If inoutState Then
            If State_LOCK <> strReserved Then GoTo errThrow
        Else
            If State_LOCK <> "" Then
                If strReserved <> State_LOCK Then GoTo errThrow
            End If
            inoutState = False
            State_LOCK = strReserved
        End If
    End If
    Exit Sub
errThrow:
    throw InvalidArgumentValueException("Invalid Reserved Value.")
End Sub

'---------------------------------------------
'function Write Arguments to Template.
Public Function funcwArgs_(Template As String, ParamArray Args()) As String
    Dim cArgs() As Variant
    cArgs = Args
    funcwArgs_ = funcwArgs(Template, cArgs)
End Function
'function Read Arguments from Stream using Template.
Public Function funcrArgs_(Template As String, tStream As ITargetStream, ParamArray Args()) As String
    Dim cArgs() As Variant
    cArgs = Args
    funcrArgs_ = funcrArgs(Template, tStream, cArgs)
End Function
'---------------------------------------------
'function Write Arguments to Template.
Public Function funcwArgs(Template As String, Args()) As String
    '%d,%s,%i,%u,%l,%f,%c,%b,%o  ,\\,\n,\r,\a,\t,\c,\xFF,\255,\0377
    '%o:object
End Function
'function Read Arguments from Stream using Template.
Public Function funcrArgs(Template As String, tStream As ITargetStream, Args()) As String
    '%d,%s,%i,%u,%l,%f,%c,%b,%o  ,\\,\n,\r,\a,\t,\c,\xFF,\255,\0377
End Function


Public Function Array_String(ParamArray str() As Variant) As String()
    Dim i As Long, strSize As Long, retVal() As String
    On Error GoTo zeroLength
    strSize = UBound(str) - LBound(str) + 1
zeroLength:
    If strSize > 0 Then
        ReDim retVal(strSize - 1)
        For i = 0 To strSize - 1
            retVal(i) = CStr(str(i))
        Next
    End If
    Array_String = retVal
End Function
Public Function Array_Object(ParamArray Objects() As Variant) As Object()
    Dim i As Long, objSize As Long, retVal() As Object
    On Error GoTo zeroLength
    objSize = UBound(Objects) - LBound(Objects) + 1
zeroLength:
    If objSize > 0 Then
        ReDim retVal(objSize - 1)
        For i = 0 To objSize - 1
            Set retVal(i) = Objects(i)
        Next
    End If
    Array_Object = retVal
End Function
Public Function Array_Double(ParamArray Doubles() As Variant) As Double()
    Dim i As Long, dblSize As Long, retVal() As Double
    On Error GoTo zeroLength
    dblSize = UBound(Doubles) - LBound(Doubles) + 1
zeroLength:
    If dblSize > 0 Then
        ReDim retVal(dblSize - 1)
        For i = 0 To dblSize - 1
            retVal(i) = CDbl(Doubles(i))
        Next
    End If
    Array_Double = retVal
End Function
Public Function Array_Single(ParamArray Singles() As Variant) As Single()
    Dim i As Long, sngSize As Long, retVal() As Single
    On Error GoTo zeroLength
    sngSize = UBound(Singles) - LBound(Singles) + 1
zeroLength:
    If sngSize > 0 Then
        ReDim retVal(sngSize - 1)
        For i = 0 To sngSize - 1
            retVal(i) = CSng(Singles(i))
        Next
    End If
    Array_Single = retVal
End Function
Public Function Array_Long(ParamArray Longs() As Variant) As Long()
    Dim i As Long, lngSize As Long, retVal() As Long
    On Error GoTo zeroLength
    lngSize = UBound(Longs) - LBound(Longs) + 1
zeroLength:
    If lngSize > 0 Then
        ReDim retVal(lngSize - 1)
        For i = 0 To lngSize - 1
            retVal(i) = CLng(Longs(i))
        Next
    End If
    Array_Long = retVal
End Function
Public Function Array_Integer(ParamArray Ints() As Variant) As Integer()
    Dim i As Long, intSize As Long, retVal() As Integer
    On Error GoTo zeroLength
    intSize = UBound(Ints) - LBound(Ints) + 1
zeroLength:
    If intSize > 0 Then
        ReDim retVal(intSize - 1)
        For i = 0 To intSize - 1
            retVal(i) = CLng(Ints(i))
        Next
    End If
    Array_Integer = retVal
End Function
Public Function Array_Byte(ParamArray Bytes() As Variant) As Byte()
    Dim i As Long, btSize As Long, retVal() As Byte
    On Error GoTo zeroLength
    btSize = UBound(Bytes) - LBound(Bytes) + 1
zeroLength:
    If btSize > 0 Then
        ReDim retVal(btSize - 1)
        For i = 0 To btSize - 1
            retVal(i) = CByte(Bytes(i))
        Next
    End If
    Array_Byte = retVal()
End Function
