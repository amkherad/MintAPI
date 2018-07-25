Attribute VB_Name = "mint_invokers"
Option Explicit

Public Enum MINT_INVOKERS_ATTRIBUTES
    MIA_NONE = 0
    MIA_PUSHOBJECT = 1
End Enum

Public Function mint_invokers_invoke( _
    ByVal Object As Object, ByVal Method As Method, _
    ByVal Prototype As MethodPrototype, ByVal Args As ArgumentList, _
    ByVal Attrs As MINT_INVOKERS_ATTRIBUTES) As Variant
'----------------
    Select Case Prototype.CallingConvension
        Case ccStdCall
            Call MoveVariantX(mint_invokers_invoke, _
                mint_invokers_invoke_STDCALL(Object, Method, Prototype, Args, Attrs))
        Case ccCdecl
            Call MoveVariantX(mint_invokers_invoke, _
                mint_invokers_invoke_CDECL(Object, Method, Prototype, Args, Attrs))
        Case Else
            Call MoveVariantX(mint_invokers_invoke, _
            mint_invokers_invoke_USER(Object, Method, Prototype, Args, Attrs))
    End Select
End Function

Private Function mint_invokers_invoke_STDCALL( _
    ByVal Object As Object, ByVal Method As Method, ByVal Prototype As MethodPrototype, _
    ByVal Args As ArgumentList, ByVal Attrs As MINT_INVOKERS_ATTRIBUTES) As Variant
'----------------
    Dim Scheme As ParametersScheme
    Set Scheme = Prototype.ParametersScheme
    If Scheme Is Nothing Then
        If Not Args Is Nothing Then _
            Set Scheme = Args.Scheme
    Else
        If Not Scheme.Validate(Args) Then _
            If Not Prototype.AllowUnsafeCall Then _
                throw Exps.InvalidOperationException("The method does not allow unsafe call.")
    End If
    
    Dim BufArgs As ArgumentList
    Set BufArgs = Args
    If Not Args Is Nothing Then
        Set Args = Args.PackForPush(Scheme) '##IMPORTANT to 1)create a lock 2)reverse 3)true byval or byref
        Dim V As Variant, Arg As Argument
        For Each V In Args
            Set Arg = V
            If Arg.IsByRef Then
                MsgBox "ByRefPush"
                Call ThreadStack.PushLong(Arg.Reference)
            Else
                If Arg.IsString Then
                    MsgBox "ByRefPush"
                    Call ThreadStack.PushLong(Arg.AbsoluteReference)
                ElseIf Arg.IsArray Then
                    MsgBox "IsArray"
                    Call ThreadStack.PushLong(Arg.AbsoluteReference)
                ElseIf Arg.IsRecord Then
                    MsgBox "IsRecord"
                    Call ThreadStack.PushLong(Arg.AbsoluteReference)
                Else
                    Dim VT As VariantTypes
                    VT = Arg.ArgumentType
                    MsgBox "Else"
                    If VT = VT_Double Then
                        Dim DblVal As Double
                        MsgBox "Double"
                        Call memcpy(DblVal, ByVal Arg.AbsoluteReference, VLEN_DOUBLE)
                        Call ThreadStack.PushDouble(DblVal)
                    Else
                        Select Case Arg.AbsoluteSize
                            Case VLEN_LONG, VLEN_INTEGER, VLEN_BYTE
                                Dim LngVal As Long
                                MsgBox "Long"
                                Call memcpy(LngVal, ByVal Arg.AbsoluteReference, VLEN_LONG)
                                Call ThreadStack.PushLong(LngVal)
                            Case VLEN_CURRENCY
                                Dim CurVal As Currency
                                MsgBox "Currency"
                                Call memcpy(CurVal, ByVal Arg.AbsoluteReference, VLEN_CURRENCY)
                                Call ThreadStack.PushCurrency(CurVal)
                            Case Else
                                throw Exps.NotSupportedException
                        End Select
                    End If
                End If
            End If
        Next
    End If
    
    If (Attrs And MIA_PUSHOBJECT) = MIA_PUSHOBJECT Then _
        Call ThreadStack.PushObject(Object)
    
    Dim Ref As Long, ReturnValueType As VariantTypes
    Ref = Method.Reference
    ReturnValueType = Prototype.ReturnValueType
    
    Select Case ReturnValueType
        Case VT_EMPTY, VT_VOID
            Call mHelper.Call(Ref)
        Case VT_Double
            mint_invokers_invoke_STDCALL = mHelper.CallDbl(Ref)
        Case VT_Currency
            mint_invokers_invoke_STDCALL = mHelper.CallInt64(Ref)
        Case Else
            mint_invokers_invoke_STDCALL = mHelper.CallInt32(Ref)
            VariantType(mint_invokers_invoke_STDCALL) = ReturnValueType
    End Select
End Function
Private Function mint_invokers_invoke_CDECL( _
    ByVal Object As Object, ByVal Method As Method, ByVal Prototype As MethodPrototype, _
    ByVal Args As ArgumentList, ByVal Attrs As MINT_INVOKERS_ATTRIBUTES) As Variant
'----------------
    throw Exps.NotImplementedException
End Function
Private Function mint_invokers_invoke_FASTCALL( _
    ByVal Object As Object, ByVal Method As Method, ByVal Prototype As MethodPrototype, _
    ByVal Args As ArgumentList, ByVal Attrs As MINT_INVOKERS_ATTRIBUTES) As Variant
'----------------
    throw Exps.NotImplementedException
End Function
Private Function mint_invokers_invoke_THISCALL( _
    ByVal Object As Object, ByVal Method As Method, ByVal Prototype As MethodPrototype, _
    ByVal Args As ArgumentList, ByVal Attrs As MINT_INVOKERS_ATTRIBUTES) As Variant
'----------------
    throw Exps.NotImplementedException
End Function
Private Function mint_invokers_invoke_PASCAL( _
    ByVal Object As Object, ByVal Method As Method, ByVal Prototype As MethodPrototype, _
    ByVal Args As ArgumentList, ByVal Attrs As MINT_INVOKERS_ATTRIBUTES) As Variant
'----------------
    throw Exps.NotImplementedException
End Function
Private Function mint_invokers_invoke_PARAMS( _
    ByVal Object As Object, ByVal Method As Method, ByVal Prototype As MethodPrototype, _
    ByVal Args As ArgumentList, ByVal Attrs As MINT_INVOKERS_ATTRIBUTES) As Variant
'----------------
    throw Exps.NotImplementedException
End Function

Private Function mint_invokers_invoke_USER( _
    ByVal Object As Object, ByVal Method As Method, ByVal Prototype As MethodPrototype, _
    ByVal Args As ArgumentList, ByVal Attrs As MINT_INVOKERS_ATTRIBUTES) As Variant
'----------------
    
    throw Exps.InvalidCallException("Not any user invoker plugin is registered.")
End Function
