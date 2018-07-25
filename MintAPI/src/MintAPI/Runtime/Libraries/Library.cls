VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Library"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"MetaObject"
'@PROJECT_LICENSE

''<class name="Library" tags="">
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

Option Base 0
Option Explicit

Implements IObject
'Implements ICloneable

'=============================================
'=============================================
'=============================================
'<section Meta Data>
'
Const CLASSID As String = "Library" '#DO NOT REMOVE IT
Const CLASSALIAS As String = APPLICATIONDOMAIN & "." & CLASSID '#DO NOT REMOVE IT
Const CLASSVERSIONSTRING As String = "0.1" '#DO NOT REMOVE IT
Const CLASSALIASSTRING As String = CLASSALIAS & "/" & CLASSVERSIONSTRING '#DO NOT REMOVE IT

Const SIGNAL_TERMINATING                As String = "terminating"
Const SLOT_METAOBJECT                   As String = "metaobject"
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
'<section Variables/Events>
'
Dim This_Constructed As Boolean
Dim myMetaObject As MetaObject

Dim p_Handle As Long
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
Friend Sub Constructor0(ByVal LibraryHandle As Long)
    p_Handle = LibraryHandle
    
    This_Constructed = True
End Sub
'Friend Sub ConstructorCopy(ByVal That As {My Object Type})
'
'    This_Constructed = True
'End Sub
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

Public Property Get FileName() As String
THROW_IfClosed
    FileName = String$(MAX_PATH_W, 0)
    Dim Result As Long
    Result = API_GetModuleFileNameUnicode(p_Handle, FileName, MAX_PATH_W)
    If Result = vbNullPtr Then throw Exps.IfError
    FileName = Left$(FileName, Result)
End Property

Public Property Get BaseAddress() As Long
THROW_IfClosed
    Dim ModInfo As API_MODULEINFO
    If API_GetModuleInformation(vbNullPtr, p_Handle, ModInfo, Len(ModInfo)) = vbNullPtr Then _
        throw Exps.IfError
    
    BaseAddress = ModInfo.lpBaseOfDll
End Property
Public Property Get EntryPointAddress() As Long
THROW_IfClosed
    Dim ModInfo As API_MODULEINFO
    If API_GetModuleInformation(vbNullPtr, p_Handle, ModInfo, Len(ModInfo)) = vbNullPtr Then _
        throw Exps.IfError
    
    EntryPointAddress = ModInfo.EntryPoint
End Property
Public Property Get ModuleMemorySize() As Long
THROW_IfClosed
    Dim ModInfo As API_MODULEINFO
    If API_GetModuleInformation(vbNullPtr, p_Handle, ModInfo, Len(ModInfo)) = vbNullPtr Then _
        throw Exps.IfError
    
    ModuleMemorySize = ModInfo.SizeOfImage
End Property

Public Property Get ModuleName() As String
THROW_IfClosed
    ModuleName = String$(MAX_PATH_W, 0)
    Dim Result As Long
    Result = API_GetModuleBaseName(vbNullPtr, p_Handle, ModuleName, MAX_PATH_W)
    If Result = vbNullPtr Then throw Exps.IfError
    ModuleName = Left$(ModuleName, Result)
End Property

Public Property Get VersionInfo() As Version 'FileVersionInfo
THROW_IfClosed
    Dim ModInfo As API_MODULEINFO
    If API_GetModuleInformation(vbNullPtr, p_Handle, ModInfo, Len(ModInfo)) = vbNullPtr Then _
        throw Exps.IfError
    
    'ModInfo.
End Property

Public Function LoadSymbol(ByVal Name As String) As Long
THROW_IfClosed
    Dim pHndl As Long
    pHndl = API_GetProcAddress(p_Handle, Name)
    If pHndl = vbNullPtr Then throw Exps.IfError
    
    LoadSymbol = pHndl
End Function

Public Function LoadMethod(ByVal Name As String) As Method ', Optional ByVal CallingConvension As MethodCallingConvension = Nothing
THROW_IfClosed
    Dim pHndl As Long
    pHndl = API_GetProcAddress(p_Handle, Name)
    If pHndl = vbNullPtr Then throw Exps.IfError
    
    Set LoadMethod = Method.FromReferenceUnsafe(Name, pHndl, ccAPICall)
    Call GC.AddToFamilyChain(LoadMethod, Me)
End Function

Public Sub Dispose()
    If p_Handle <> vbNullPtr Then
        If API_FreeLibrary(p_Handle) = vbNullPtr Then _
            throw Exps.IfError
    End If
End Sub
    

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
    Equals = Objects.Equals(Me, Value)
End Function
Public Function ToString() As String
    ToString = CLASSALIAS
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

'Public Function Clone() As Object
'
'End Function
'
'</section>
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section Private helpers>
'
'Private Sub Class_Initialize()
'    ...
'End Sub
Private Sub Class_Terminate()
    If Not myMetaObject Is Nothing Then
        Signal.Emit Me, SIGNAL_TERMINATING
        Call myMetaObject.Dispose
        Set myMetaObject = Nothing
    End If
    
    Call Dispose
End Sub

Private Sub THROW_IfClosed()
    If p_Handle = vbNullPtr Then _
        throw Exps.InvalidOperationException
End Sub
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
'Private Function ICloneable_Clone() As Object: Set ICloneable_Clone = Clone: End Function
'
'</section>
'---------------------------------------------
'---------------------------------------------
