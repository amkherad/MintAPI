Attribute VB_Name = "mint_functiondelegator"
'    Copyright (c) 2004 Kelly Ethridge
'
'    This file is part of VBCorLib.
'
'    VBCorLib is free software; you can redistribute it and/or modify
'    it under the terms of the GNU Library General Public License as published by
'    the Free Software Foundation; either version 2.1 of the License, or
'    (at your option) any later version.
'
'    VBCorLib is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU Library General Public License for more details.
'
'    You should have received a copy of the GNU Library General Public License
'    along with Foobar; if not, write to the Free Software
'    Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
'    Module: modFunctionDelegator
'
Option Base 0
Option Explicit

Private Const DELEGATE_ASM As Currency = -368956918007638.6215@     ' from Matt Curland

Public Const PAGE_NOACCESS = 1&
Public Const PAGE_READONLY = 2&
Public Const PAGE_READWRITE = 4&
Public Const PAGE_WRITECOPY = 8&
Public Const PAGE_EXECUTE = &H10&
Public Const PAGE_EXECUTE_READ = &H20&
Public Const PAGE_EXECUTE_READWRITE = &H40&
Public Const PAGE_EXECUTE_WRITECOPY = &H80&
Public Const PAGE_GUARD = &H100&
Public Const PAGE_NOCACHE = &H200&

Private Const ASM_SYN_RET   As Long = &HC3 '195
Private Const ASM_SYN_HLT   As Long = &HF4 '244

Public Type FunctionDelegator
    pVTable As Long
    pfn As Long
    cRefs As Long
    Func(3) As Long
End Type

Private mDelegateASM As Currency
Private mAsm As Long

Private mInitDelegatorQueryInterface    As Long
Private mInitDelegatorAddRelease        As Long
Private mNewDelegatorQueryInterface     As Long
Private mNewDelegatorAddRef             As Long
Private mNewDelegatorRelease            As Long


Public Function InitDelegator(ByRef Delegator As FunctionDelegator, Optional ByVal pfn As Long = 0) As IUnknown
    If mAsm = 0 Then Init
    
    With Delegator
        .pfn = pfn
        .pVTable = VarPtr(.Func(0))
        .Func(0) = mInitDelegatorQueryInterface
        .Func(1) = mInitDelegatorAddRelease
        .Func(2) = mInitDelegatorAddRelease
        .Func(3) = mAsm
        .pfn = pfn
    End With
    
    ObjectPtr(InitDelegator) = VarPtr(Delegator)
End Function


Public Function CreateDelegator(ByVal pfn As Long) As IUnknown
    Dim This As Long
    Dim Struct As FunctionDelegator
    
    If mAsm = 0 Then Init

    This = API_CoTaskMemAlloc(Len(Struct))
    If This = 0 Then throw Exps.OutOfMemoryException
    
    With Struct
        .pVTable = This + 12
        .Func(0) = mNewDelegatorQueryInterface
        .Func(1) = mNewDelegatorAddRef
        .Func(2) = mNewDelegatorRelease
        .Func(3) = mAsm
        .pfn = pfn
        .cRefs = 1
    End With

    Call memcpy(ByVal This, Struct, Len(Struct))
    ObjectPtr(CreateDelegator) = This
End Function

Public Function GetAddressOf(ByVal pfnAddress As Long) As Long
    GetAddressOf = pfnAddress
End Function

Private Sub Init()
    mDelegateASM = DELEGATE_ASM
    mAsm = VarPtr(mDelegateASM)
    
    Call API_VirtualProtect(mDelegateASM, 8, PAGE_EXECUTE_READWRITE, 0&)
    
    mInitDelegatorQueryInterface = GetAddressOf(AddressOf InitDelegator_QueryInterface)
    mInitDelegatorAddRelease = GetAddressOf(AddressOf InitDelegator_AddRefRelease)
    mNewDelegatorQueryInterface = GetAddressOf(AddressOf NewDelegator_QueryInterface)
    mNewDelegatorAddRef = GetAddressOf(AddressOf NewDelegator_AddRef)
    mNewDelegatorRelease = GetAddressOf(AddressOf NewDelegator_Release)
End Sub

Private Function InitDelegator_QueryInterface(ByVal This As Long, ByVal riid As Long, ByRef pvObj As Long) As Long
    pvObj = This
End Function

Private Function InitDelegator_AddRefRelease(ByVal This As Long) As Long
    ' do nothing
End Function


'Public Function GetMethodBody(ByVal Method As Method) As ByteArray
'    If Method Is Nothing Then _
'        throw Exps.ArgumentNullException("Method is null.")
'
'    If Not Method.Executable Then _
'        throw Exps.InvalidStatusException("Target method is not executable.")
'
'    Dim IntPtr As Long, BSyntax As Byte
'    IntPtr = Method.Reference
'
'    Dim BA As New ByteArray
'    BA.ExtraBuffer = 200 'KB
'
'    Call memcpy(BSyntax, ByVal IntPtr, 1)
'    Dim stpSyn As Long
'    While (BSyntax <> ASM_SYN_RET)
'        Call BA.Append(BSyntax)
'        stpSyn = stpSyn + 1
'        Call memcpy(BSyntax, ByVal IntPtr, 1)
'    Wend
'
'    Set GetMethodBody = BA
'End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   VTable functions used by a newly created lightweight COM function delegator
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function NewDelegator_QueryInterface(ByRef This As FunctionDelegator, ByVal riid As Long, ByRef pvObj As Long) As Long
    pvObj = VarPtr(This)
    Call NewDelegator_AddRef(This)
End Function
Private Function NewDelegator_AddRef(ByRef This As FunctionDelegator) As Long
    With This
        .cRefs = .cRefs + 1
        NewDelegator_AddRef = .cRefs
    End With
End Function
Private Function NewDelegator_Release(ByRef This As FunctionDelegator) As Long
    With This
        .cRefs = .cRefs - 1
        NewDelegator_Release = .cRefs
        If .cRefs = 0 Then _
            Call API_CoTaskMemFree(VarPtr(This))
    End With
End Function

