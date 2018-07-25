VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GuidStatic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'@PROJECT_LICENSE

''<class name="GuidStatic" static tags="">
''  <summary>
''    Some class description.
''  </summary>
''  <using>Some class alias.</using>
''  <idea>Me</idea>
''  <sample></sample>
''</class>

Option Base 0
Option Explicit

'=============================================
'=============================================
'=============================================
'<section Meta Data>
'
Const CLASSID As String = "GuidStatic" '#DO NOT REMOVE IT
Const CLASSALIAS As String = APPLICATIONDOMAIN & "." & CLASSID '#DO NOT REMOVE IT
Const CLASSVERSIONSTRING As String = "0.1" '#DO NOT REMOVE IT
Const CLASSALIASSTRING As String = CLASSALIAS & "/" & CLASSVERSIONSTRING '#DO NOT REMOVE IT

'
'</section>
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section API/Type Declarations>
'

'
'</section>
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section Creators>
'
Public Function Create() As Guid
Attribute Create.VB_UserMemId = 0
    Set Create = New Guid
    Dim StdG As StdGuid
    Call memcpy(StdG, API_CoCreateGuid, SIZEOF_GUID)
    Call Create.Constructor1(StdG)
End Function

Public Function FromStdGuid(ByRef StdGuid As StdGuid) As Guid
    Set FromStdGuid = New Guid
    Call FromStdGuid.Constructor1(StdGuid)
End Function

Public Function FromBytes(ByRef Bytes() As Byte) As Guid
    Set FromBytes = New Guid
    Call FromBytes.Constructor2(Bytes)
End Function

Public Function FromByteArray(ByVal ByteArray As ByteArray) As Guid
    Set FromByteArray = New Guid
    Call FromByteArray.Constructor2(ByteArray.SubArray(0, SIZEOF_GUID).ToConstData)
End Function

Public Function FromMemory(ByVal MemoryBlock As IMemoryBlock) As Guid
    Set FromMemory = New Guid
    Call FromMemory.Constructor6(MemoryBlock)
End Function
Friend Function FromReference(ByVal Ref As Long) As Guid
    Set FromReference = New Guid
    Call FromReference.Constructor8(Ref)
End Function

Public Function FromParts(ByVal P1 As Long, ByVal P2 As Integer, ByVal P3 As Integer, ByRef P4() As Byte) As Guid
    Set FromParts = New Guid
    Call FromParts.Constructor3(P1, P2, P3, P4)
End Function

Public Function FromPartsExt(ByVal P1 As Long, ByVal P2 As Integer, ByVal P3 As Integer, ByVal P4_1 As Byte, ByVal P4_2 As Byte, ByVal P4_3 As Byte, ByVal P4_4 As Byte, ByVal P4_5 As Byte, ByVal P4_6 As Byte, ByVal P4_7 As Byte, ByVal P4_8 As Byte) As Guid
    Set FromPartsExt = New Guid
    Call FromPartsExt.Constructor3(P1, P2, P3, Arrays.Bytes(P4_1, P4_2, P4_3, P4_4, P4_5, P4_6, P4_7, P4_8))
End Function

Public Function FromProgramID(ByVal ProgID As String) As Guid
    Set FromProgramID = New Guid
    Dim APIStdGuid As API_StdGuid, StdGuid As StdGuid
    APIStdGuid = API_CLSIDFromProgID(ProgID)
    Call memcpy(StdGuid, APIStdGuid, SIZEOF_GUID)
    Call FromProgramID.Constructor1(StdGuid)
End Function

Public Function Parse(ByVal StrGuid As String) As Guid
    Set Parse = New Guid
    Call Parse.Constructor0(StrGuid)
End Function

Friend Function FromAPIGuid(ByRef Value As API_StdGuid) As Guid
    Set FromAPIGuid = New Guid
    Call FromAPIGuid.Constructor7(Value)
End Function
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

Public Property Get EmptyGuid() As Guid
    Set EmptyGuid = New Guid
End Property

Public Property Get IUnknown() As Guid
    Set IUnknown = Parse(GUIDS_IUnknown)
End Property
Public Property Get IDispatch() As Guid
    Set IDispatch = Parse(GUIDS_IDispatch)
End Property

Friend Property Get ITypeInfo() As Guid
    Set ITypeInfo = Parse(GUIDS_ITypeInfo)
End Property
Friend Property Get ITypeInfo2() As Guid
    Set ITypeInfo2 = Parse(GUIDS_ITypeInfo2)
End Property
Friend Property Get ITypeLib() As Guid
    Set ITypeLib = Parse(GUIDS_ITypeLib)
End Property
Friend Property Get ITypeLib2() As Guid
    Set ITypeLib2 = Parse(GUIDS_ITypeLib2)
End Property

'*********************************************
'*********************************************
'
'</section>
'---------------------------------------------
'---------------------------------------------
