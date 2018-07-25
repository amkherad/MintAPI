VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Guid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"MetaObject"
'@PROJECT_LICENSE

''<class name="Guid" tags="immutable">
''  <summary>
''    Some class description.
''  </summary>
''  <constructors>
''      <constructor name="" friend>
''          Some constructor description.
''      </constructor>
''  </constructors>
''  <using>Some class alias.</using>
''  <idea>Me</idea>
''  <sample></sample>
''</class>

'' ##This class written by Kelly Ethridge - VBCorLib - LGPLv2

Option Base 0
Option Explicit

Implements IObject

'=============================================
'=============================================
'=============================================
'<section Meta Data>
'
Const CLASSID As String = "Guid" '#DO NOT REMOVE IT
Const CLASSALIAS As String = APPLICATIONDOMAIN & "." & CLASSID '#DO NOT REMOVE IT
Const CLASSVERSIONSTRING As String = "0.1" '#DO NOT REMOVE IT
Const CLASSALIASSTRING As String = CLASSALIAS & "/" & CLASSVERSIONSTRING '#DO NOT REMOVE IT

Const SIGNAL_TERMINATING As String = "terminating"
Const SLOT_METAOBJECT As String = "metaobject"
'
'</section>
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section API/Type Declarations>
'
Public Type StdGuid
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type
'
'</section>
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section Variables/Events>
'
Dim This_Constructed As Boolean
Dim myMetaObject As MetaObject

Dim InnerGuid As StdGuid
'
'</section>
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section Constructors>
'
''<summary>
''  Basic constructor to use by MintAPI class provider,
''  When creating new instance of this class.
''</summary>
''<params>
''  <param name="Arguments">
''      A list of construction arguments.
''  </param>
''</params>
''<cargs order="0" count="1">
''  <carg type="long" name="ID"></carg>
''</cargs>
Friend Sub Constructor(ByVal CInfo As ObjectConstructionInfo)
    
    This_Constructed = True
End Sub
Friend Sub Constructor0(ByVal StrGuid As String)
    Dim StdG As API_StdGuid, StrGLen As Long
    On Error GoTo CatchErr
    StrGLen = Len(StrGuid)
    If StrGLen = 38 Then
        If Left$(StrGuid, 1) <> "{" Then Mid$(StrGuid, 1, 1) = "{"
        If Right$(StrGuid, 1) <> "}" Then Mid$(StrGuid, StrGLen, 1) = "}"
    ElseIf StrGLen = 36 Then
        StrGuid = "{" & StrGuid & "}"
    Else
        GoTo CatchErr
    End If
    StdG = API_GUIDFromString(StrGuid)
    Call memcpy(InnerGuid, StdG, SIZEOF_GUID)
    
    This_Constructed = True
    Exit Sub
CatchErr:
    Throw_InvalidFormat
End Sub
Friend Sub Constructor1(ByRef StdGuid As StdGuid)
    InnerGuid = StdGuid
    
    This_Constructed = True
End Sub
Friend Sub Constructor2(ByRef Bytes() As Byte)
    If Arrays.IsNull(Bytes) Then _
        throw Exps.ArgumentNullException
    If Arrays.Length(Bytes) <> SIZEOF_GUID Then _
        throw Exps.InvalidArgumentException("Array must be " & SIZEOF_GUID & " bytes in length.")
    
    Call memcpy(InnerGuid, ByVal Arrays.GetDataPointerOf(Bytes), SIZEOF_GUID)
    
    This_Constructed = True
End Sub
Friend Sub Constructor3(ByVal P1 As Long, ByVal P2 As Integer, ByVal P3 As Integer, ByRef P4() As Byte)
    With InnerGuid
    
        If Arrays.IsNull(P4) Then _
            throw Exps.ArgumentNullException("P4")
        If Arrays.Length(P4) <> 8 Then _
            throw Exps.InvalidArgumentException("Array must be 8 bytes in length.")
    
        .Data1 = P1
        .Data2 = P2
        .Data3 = P3
        
        Call memcpy(.Data4(0), ByVal Arrays.GetDataPointerOf(P4), 8)
    End With
    
    This_Constructed = True
End Sub
Friend Sub Constructor4(ByVal ProgID As String)
    Dim StdG As API_StdGuid
    On Error GoTo CatchErr
    StdG = API_CLSIDFromProgID(ProgID)
    Call memcpy(InnerGuid, StdG, SIZEOF_GUID)
    
    This_Constructed = True
    Exit Sub
CatchErr:
    Throw_InvalidFormat
End Sub
Friend Sub Constructor5()
    
    This_Constructed = True
End Sub
Friend Sub Constructor6(ByVal MemoryBlock As IMemoryBlock)
    If MemoryBlock Is Nothing Then throw Exps.ArgumentNullException("MemoryBlock")
    If MemoryBlock.Limit <> SIZEOF_GUID Then throw Exps.InvalidOperationException
    
    Call memcpy(InnerGuid, ByVal MemoryBlock.Reference, SIZEOF_GUID)
    
    This_Constructed = True
End Sub
Friend Sub Constructor7(ByRef StdGuid As API_StdGuid)
    Call memcpy(InnerGuid, StdGuid, SIZEOF_GUID)
    
    This_Constructed = True
End Sub
Friend Sub Constructor8(ByVal Ref As Long)
    Call memcpy(InnerGuid, ByVal Ref, SIZEOF_GUID)
    
    This_Constructed = True
End Sub
Friend Sub ConstructorCopy(ByVal Guid As Guid)
    If Guid Is Nothing Then throw Exps.ArgumentNullException("Guid")
    Call memcpy(InnerGuid, Guid.ToStdGuid, SIZEOF_GUID)
    
    This_Constructed = True
End Sub
'
'</section>
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section Class Public Body Members>
'
'*********************************************
'*********************************************


Friend Property Get Handle() As Long
    Handle = VarPtr(InnerGuid)
End Property


Public Function ToStdGuid() As StdGuid
    ToStdGuid = InnerGuid
End Function

Public Function ToByteArray() As ByteArray
    Dim GuidBytes(SIZEOF_GUID) As Byte
    Call memcpy(GuidBytes(0), InnerGuid, SIZEOF_GUID)
    Set ToByteArray = ByteArray(GuidBytes)
End Function

Public Function CompareTo(ByVal Guid As Guid) As CompareResults
    If Guid Is Nothing Then throw Exps.ArgumentNullException("Guid")
    CompareTo = Compare(InnerGuid, Guid.ToStdGuid)
End Function

Public Function ToProgramID() As String
    Dim StdG As API_StdGuid, outProgID As Long, RetVal As Long
    Call memcpy(StdG, InnerGuid, SIZEOF_GUID)
    RetVal = API_ProgIDFromCLSID(StdG, outProgID)
    Select Case RetVal
        Case S_OK
            'Do nothing
        Case ERROR_REGDB_E_CLASSNOTREG
            throw Exps.InvalidOperationException("The Program ID is not a registered class.")
        Case ERROR_REGDB_E_READREGDB
            throw Exps.InvalidOperationException("The Registry could not be read.")
        Case Else
            throw Exps.InvalidOperationException
    End Select
    
    ToProgramID = API_SysAllocString(outProgID)
    Call API_CoTaskMemFree(outProgID)
End Function

'*********************************************
'*********************************************
'
'</section>
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section Basic Members>
'
Public Function Equals(Value As Variant) As Boolean
    If IsObject(Value) Then
        If Not Value Is Nothing Then
            If TypeOf Value Is Guid Then
                Dim Guid As Guid
                Set Guid = Value
                Equals = (Compare(InnerGuid, Guid.ToStdGuid) = crEqual)
            End If
        End If
    Else
        Equals = Objects.Equals(Me, Value)
    End If
End Function
Public Function ToString(Optional ByVal Format As String) As String
    'ToString = CLASSALIAS
    Dim RetVal As String
    Dim RetSize As Long
    Dim Gu As API_StdGuid
    
    Call memcpy(Gu, InnerGuid, SIZEOF_GUID)
    
    RetVal = String$(TINYLPSTR, 0)
    RetSize = API_StringFromGUID2(Gu, RetVal, TINYLPSTR)
    ToString = Left$(RetVal, RetSize - 1)
    
    Select Case UCase$(Format)
        Case "D"
            ToString = Mid$(ToString, 2, 36)
        Case "N"
            ToString = Replace$(Mid$(ToString, 2, 36), "-", "")
        Case "P"
            Mid$(ToString, 1, 1) = "("
            Mid$(ToString, 38, 1) = ")"
        Case "B", ""
            ' already formatted to this and is the default
        Case Else
            Throw_InvalidFormat
    End Select
End Function
Public Function GetHashCode() As Long
    GetHashCode = Objects.GetHashCode(Me)
End Function
Public Function MetaObject() As MetaObject
    If myMetaObject Is Nothing Then
        Call Objects.MetaObject.Setup(myMetaObject, Me, True, CLASSALIASSTRING)
    End If
    Set MetaObject = myMetaObject
End Function

Public Function Clone() As Guid
    Set Clone = New Guid
    Call Clone.ConstructorCopy(Me)
End Function
'
'</section>
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section Private helpers>
'
Private Sub Class_Terminate()
    If Not myMetaObject Is Nothing Then
        Signal.Emit Me, SIGNAL_TERMINATING
        Call myMetaObject.Dispose
        Set myMetaObject = Nothing
    End If
End Sub

Private Sub Throw_InvalidFormat()
    throw Exps.InvalidOperationException("Invalid Guid format.")
End Sub

Friend Function ToAPIStdGuid() As API_StdGuid
    Call memcpy(ToAPIStdGuid, InnerGuid, SIZEOF_GUID)
End Function

Private Function CTT(ByVal lng1 As Long, ByVal lng2 As Long) As CompareResults
'CompareTheseTwo
    If lng1 > lng2 Then
        CTT = crGreater
    ElseIf lng1 < lng2 Then
        CTT = crLess
    Else
        CTT = crEqual
    End If
End Function
Private Function Compare(G1 As StdGuid, G2 As StdGuid) As CompareResults
    Compare = CTT(G1.Data1, G2.Data1): If Compare <> crEqual Then Exit Function
    Compare = CTT(G1.Data2, G2.Data2): If Compare <> crEqual Then Exit Function
    Compare = CTT(G1.Data3, G2.Data3): If Compare <> crEqual Then Exit Function
    
    Dim i As Long
    For i = 0 To 7
        Compare = CTT(G1.Data4(i), G2.Data4(i)): If Compare <> crEqual Then Exit Function
    Next
End Function
'
'</section>
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section IObject Implementation>
'
Private Sub IObject_Construct(ByVal CInfo As ObjectConstructionInfo): Call Constructor(CInfo): End Sub
Private Function IObject_Equals(Value As Variant) As Boolean: IObject_Equals = Equals(Value): End Function
Private Function IObject_GetHashCode() As Long: IObject_GetHashCode = GetHashCode: End Function
Private Function IObject_MetaObject() As MetaObject: Set IObject_MetaObject = MetaObject: End Function
Private Function IObject_ToString() As String: IObject_ToString = ToString: End Function
Private Function IObject_Stub(ByVal ProxyInfo As ObjectProxyInfo) As Boolean
    
    IObject_Stub = False
End Function
'
'</section>
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section ICloneable Implementations>
'
Private Function ICloneable_Clone() As Object: Set ICloneable_Clone = Clone: End Function
'
'</section>
'---------------------------------------------
'---------------------------------------------
