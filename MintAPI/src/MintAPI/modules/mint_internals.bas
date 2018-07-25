Attribute VB_Name = "mint_internals"
Option Explicit

'=============================================
'=============================================
'=============================================
'<section DynamicObject Methods>
'
'=========================
Public Function mint_internals_DynamicObject_QueryInterface( _
    ByVal This As DynamicObject, _
    ByRef riid As StdGuid, _
    ByRef pvObj As Long) As Long
    '-----------------------------
    mint_internals_DynamicObject_QueryInterface = _
        This.QueryInterface(riid, pvObj)
End Function
'=========================
Public Function mint_internals_DynamicObject_GetTypeInfo( _
    ByVal This As DynamicObject, _
    ByRef iTInfo As ITypeInfo, _
    ByRef LCID As Long) As ITypeInfo
    '-----------------------------
    Set mint_internals_DynamicObject_GetTypeInfo = _
        This.GetTypeInfo(iTInfo, LCID)
End Function
Public Function mint_internals_DynamicObject_GetIDsOfNames( _
    ByVal This As DynamicObject, _
    ByRef riid As StdGuid, _
    ByRef rgszNames() As String, _
    ByRef cNames As Long, _
    ByRef LCID As Long) As Long
    '-----------------------------
    mint_internals_DynamicObject_GetIDsOfNames = _
        This.GetIDsOfNames(riid, rgszNames, cNames, LCID)
End Function
Public Function mint_internals_DynamicObject_GetTypeInfoCount( _
    ByVal This As DynamicObject) As Long
    '-----------------------------
    mint_internals_DynamicObject_GetTypeInfoCount = _
        This.GetTypeInfoCount()
End Function
Public Function mint_internals_DynamicObject_Invoke( _
    ByVal This As DynamicObject, _
    ByRef dispIdMember As Long, _
    ByRef riid As StdGuid, _
    ByRef LCID As Long, _
    ByRef wFlags As Long, _
    ByRef pDispParams As API_DISPPARAMS, _
    ByRef pVarResult As Variant, _
    ByRef pExcepInfo As API_EXCEPINFO) As Long
    '-----------------------------
    mint_internals_DynamicObject_Invoke = _
        This.Invoke(dispIdMember, riid, LCID, wFlags, _
                pDispParams, pVarResult, pExcepInfo)
End Function
'=========================
'
'</section>
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section DynamicObject Methods>
'
'=========================
Public Function mint_internals_DynamicClassInfo_QueryInterface( _
    ByVal This As DynamicClassInfo, _
    ByRef riid As StdGuid, _
    ByRef pvObj As Long) As Long
    '-----------------------------
    mint_internals_DynamicClassInfo_QueryInterface = _
        This.QueryInterface(riid, pvObj)
End Function
'=========================
'
'</section>
'---------------------------------------------
'---------------------------------------------
